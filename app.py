"""
Zoom Participation Tracker - Streamlit Web App
A user-friendly tool to generate attendance reports from Zoom transcripts.

Author: Jorge Bienvenido Cevallos Bravo
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
import pytz
from collections import defaultdict
import unicodedata
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.piecharts import Pie
import tempfile
import os

# ============================================================================
# CONFIGURATION
# ============================================================================

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


def parse_transcript(transcript_content, teacher_name):
    """Parse the Zoom transcript and count words per speaker."""
    word_counts = defaultdict(int)
    teacher_word_count = 0
    
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
            word_count = count_words(line)
            
            # Check if it's the teacher
            if fuzzy_match_name(current_speaker, teacher_name):
                teacher_word_count += word_count
            else:
                word_counts[current_speaker] += word_count
    
    return dict(word_counts), teacher_word_count


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


def create_student_report(attendance_data, output_path, date_str, course_code, teacher_name):
    """Create a professional PDF attendance report for students."""
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
    title = Paragraph("STUDENT PARTICIPATION REPORT", title_style)
    story.append(title)
    story.append(Spacer(1, 0.2*inch))
    
    # Header information
    header_info = [
        f"<b>Course:</b> {course_code}",
        f"<b>Date:</b> {date_str}",
        f"<b>Instructor:</b> {teacher_name}"
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
        alignment=TA_LEFT,
        fontName='Helvetica'
    )
    
    # Create summary as invisible table for perfect alignment
    summary_data = [
        [Paragraph(f"<b>Total Students:</b> {total_students}", summary_style)],
        [Paragraph(f"<b>Participated:</b> {present_count}", summary_style)],
        [Paragraph(f"<b>No active participation or absent:</b> {absent_count}", summary_style)]
    ]
    
    summary_table = Table(summary_data, colWidths=[7.3*inch])
    summary_table.setStyle(TableStyle([
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(summary_table)
    story.append(Spacer(1, 0.15*inch))
    
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
         fontName='Helvetica',
         leading=14
   )
    
    table_data = [
        [
            Paragraph('<b>FULL NAME</b>', header_text_style),
            Paragraph('<b>PARTICIPATION</b>', header_text_style),
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
            participation_msg = "The student did not participate and may be marked as not present."
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
    table = Table(table_data, colWidths=[3*inch, 1.3*inch, 3*inch])
    
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
        ('VALIGN', (0, 0), (0, -1), 'MIDDLE'),
        ('VALIGN', (1, 0), (-1, -1), 'MIDDLE'),
        
        # Alternating row colors
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f3f4f6')]),
        
        # Grid
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#9ca3af')),
        
        # Padding
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 1), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
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
    ecuador_tz = pytz.timezone('America/Guayaquil')
    now_ecuador = datetime.now(ecuador_tz)
    footer = Paragraph(f"Report generated on {now_ecuador.strftime('%B %d, %Y at %I:%M %p')}", footer_style)
    story.append(footer)
    
    # Build PDF
    doc.build(story)


def create_teacher_analytics_report(teacher_word_count, student_word_count, output_path, date_str, course_code, teacher_name):
    """Create a professional PDF analytics report for teachers with pie chart."""
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
    title = Paragraph("TEACHER ANALYTICS REPORT", title_style)
    story.append(title)
    story.append(Spacer(1, 0.2*inch))
    
    # Header information
    header_info = [
        f"<b>Course:</b> {course_code}",
        f"<b>Date:</b> {date_str}",
        f"<b>Instructor:</b> {teacher_name}"
    ]
    
    for info in header_info:
        story.append(Paragraph(info, header_style))
    
    story.append(Spacer(1, 0.5*inch))
    
    # Calculate percentages
    total_words = teacher_word_count + student_word_count
    if total_words > 0:
        teacher_percentage = (teacher_word_count / total_words) * 100
        student_percentage = (student_word_count / total_words) * 100
    else:
        teacher_percentage = 0
        student_percentage = 0
    
    # Create centered pie chart without labels
    drawing = Drawing(450, 250)
    
    pie = Pie()
    pie.x = 125
    pie.y = 25
    pie.width = 200
    pie.height = 200
    
    # Data for pie chart
    pie.data = [teacher_percentage, student_percentage]
    pie.labels = None  # We'll add labels separately in a box
    
    # Colors: Orange for teacher, Grey for students
    pie.slices[0].fillColor = colors.HexColor('#f97316')  # Orange
    pie.slices[1].fillColor = colors.HexColor('#9ca3af')  # Grey
    
    # Slice styling
    pie.slices[0].strokeColor = colors.white
    pie.slices[0].strokeWidth = 2
    pie.slices[1].strokeColor = colors.white
    pie.slices[1].strokeWidth = 2
    
    drawing.add(pie)
    
    story.append(drawing)
    story.append(Spacer(1, 0.3*inch))
    
    # Create centered legend box with labels
    legend_style = ParagraphStyle(
        'Legend',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#1f2937'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        spaceAfter=4
    )
    
    # Legend data with color indicators
    legend_data = [
        [
            Paragraph('<font color="#f97316">■</font> Teacher Speaking Contribution:', legend_style),
            Paragraph(f'{teacher_percentage:.1f}%', legend_style)
        ],
        [
            Paragraph('<font color="#9ca3af">■</font> Student Speaking Contribution:', legend_style),
            Paragraph(f'{student_percentage:.1f}%', legend_style)
        ]
    ]
    
    # Create centered legend table
    legend_table = Table(legend_data, colWidths=[3.2*inch, 0.8*inch])
    legend_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
    ]))
    
    story.append(legend_table)
    story.append(Spacer(1, 0.5*inch))
    
    # Footer
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#6b7280'),
        alignment=TA_CENTER,
        fontName='Helvetica-Oblique'
    )
    ecuador_tz = pytz.timezone('America/Guayaquil')
    now_ecuador = datetime.now(ecuador_tz)
    footer = Paragraph(f"Report generated on {now_ecuador.strftime('%B %d, %Y at %I:%M %p')}", footer_style)
    story.append(footer)
    
    # Build PDF
    doc.build(story)


# ============================================================================
# STREAMLIT APP
# ============================================================================

def main():
    st.set_page_config(
        page_title="Zoom Participation Tracker",
        page_icon="📊",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    
    # Custom CSS with dark elegant theme
    st.markdown("""
        <style>
        /* Import Google Fonts */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Crimson+Pro:wght@400;600&display=swap');
        
        :root {
            --slate-dark: #1e272e;
            --slate-darker: #101417;
            --paper-bg: #ffffff;
            --accent-copper: #e67e22;
            --accent-hover: #d35400;
            --accent-green: #27ae60;
            --border-subtle: rgba(255,255,255,0.08);
            --text-dim: #7f8c8d;
            --text-medium: #57606f;
        }
        
        /* Global styles - Force dark background everywhere */
        html, body, [data-testid="stAppViewContainer"], .main, .stApp {
            background-color: var(--slate-darker) !important;
            background: var(--slate-darker) !important;
        }
        
        .main {
            padding: 0;
        }
        
        .block-container {
            padding-top: 3rem;
            padding-bottom: 3rem;
            max-width: 780px;
        }
        
        /* Override Streamlit defaults */
        [data-testid="stHeader"] {
            background-color: transparent;
        }
        
        section[data-testid="stSidebar"] {
            background-color: var(--slate-dark);
        }
        
        /* Header */
        .app-header {
            background: var(--slate-dark);
            border: 1px solid rgba(230, 126, 34, 0.2);
            border-radius: 4px;
            padding: 50px 40px;
            margin-bottom: 45px;
        }
        
        .app-title {
            font-family: 'Crimson Pro', Georgia, serif;
            font-size: 1.8rem;
            color: #ffffff !important;
            margin: 0 0 8px 0;
            text-align: center;
            font-weight: 600;
        }
        
        .app-subtitle {
            font-size: 0.7rem;
            text-transform: uppercase;
            letter-spacing: 2px;
            color: var(--accent-copper);
            text-align: center;
            font-weight: 800;
            margin: 0;
        }
        
        /* Section headers */
        .section-label {
            font-size: 0.65rem;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            color: #ffffff !important;
            font-weight: 700;
            margin: 35px 0 18px 0;
            font-family: 'Inter', sans-serif;
        }
        
        /* Input fields */
        .stTextInput > div > div > input {
            background: var(--slate-darker);
            border: 1px solid var(--border-subtle);
            border-radius: 4px;
            color: #ffffff;
            font-family: 'Inter', sans-serif;
            font-size: 0.95rem;
            padding: 13px 14px;
        }
        
        .stTextInput > div > div > input:focus {
            border-color: var(--accent-copper);
            box-shadow: none;
        }
        
        .stTextInput > label {
            font-size: 0.65rem;
            text-transform: uppercase;
            font-weight: 700;
            color: #ffffff !important;
            letter-spacing: 1.2px;
            font-family: 'Inter', sans-serif;
        }
        
        /* File uploader */
        .stFileUploader {
            background: var(--slate-dark);
            border: 1px solid var(--border-subtle);
            border-radius: 4px;
            padding: 20px;
        }
        
        .stFileUploader:hover {
            border-color: rgba(230, 126, 34, 0.3);
        }
        
        .stFileUploader > label {
            font-size: 0.65rem;
            text-transform: uppercase;
            font-weight: 700;
            color: #ffffff !important;
            letter-spacing: 1.2px;
            font-family: 'Inter', sans-serif;
        }
        
        .stFileUploader section {
            border: 2px dashed rgba(230, 126, 34, 0.3);
            border-radius: 4px;
            padding: 25px;
        }
        
        .stFileUploader section:hover {
            border-color: var(--accent-copper);
            background: rgba(230, 126, 34, 0.03);
        }
        
        /* Buttons */
        .stButton > button {
            background: var(--accent-copper);
            color: white;
            border: none;
            padding: 14px 28px;
            font-size: 0.8rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            border-radius: 4px;
            font-family: 'Inter', sans-serif;
            transition: background 0.2s;
            width: 100%;
        }
        
        .stButton > button:hover {
            background: var(--accent-hover);
        }
        
        /* Download buttons */
        .stDownloadButton > button {
            background: var(--accent-green);
            color: white;
            border: none;
            padding: 12px 24px;
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1.2px;
            border-radius: 4px;
            font-family: 'Inter', sans-serif;
            transition: background 0.2s;
            width: 100%;
        }
        
        .stDownloadButton > button:hover {
            background: #229954;
        }
        
        /* Metrics */
        div[data-testid="metric-container"] {
            background: var(--slate-dark);
            border: 1px solid var(--border-subtle);
            border-radius: 4px;
            padding: 20px;
        }
        
        div[data-testid="metric-container"] > label {
            font-size: 0.65rem;
            text-transform: uppercase;
            color: #ffffff !important;
            letter-spacing: 1.2px;
            font-weight: 700;
        }
        
        div[data-testid="metric-container"] > div {
            color: #ffffff;
            font-size: 1.8rem;
            font-weight: 700;
        }
        
        /* Expander */
        .streamlit-expanderHeader {
            background: var(--slate-dark);
            border: 1px solid var(--border-subtle);
            border-radius: 4px;
            font-size: 0.8rem;
            color: #ffffff;
            font-weight: 600;
            padding: 15px 18px;
        }
        
        .streamlit-expanderHeader:hover {
            border-color: rgba(230, 126, 34, 0.3);
        }
        
        .streamlit-expanderContent {
            background: var(--slate-dark);
            border: 1px solid var(--border-subtle);
            border-top: none;
            border-radius: 0 0 4px 4px;
            color: #ffffff;
            padding: 20px;
        }
        
        /* Info/Success/Error messages */
        .stAlert {
            background: var(--slate-dark);
            border: 1px solid var(--border-subtle);
            border-left: 4px solid var(--accent-copper);
            border-radius: 4px;
            color: #ffffff !important;
            padding: 15px 18px;
        }
        
        .stSuccess {
            border-left-color: var(--accent-green);
        }
        
        .stError {
            border-left-color: #e74c3c;
        }
        
        /* Success card */
        .success-card {
            background: var(--slate-dark);
            border: 1px solid rgba(39, 174, 96, 0.3);
            border-radius: 4px;
            padding: 28px;
            text-align: center;
            margin: 30px 0;
        }
        
        .success-card h3 {
            color: var(--accent-green);
            font-family: 'Crimson Pro', Georgia, serif;
            font-size: 1.3rem;
            margin: 0 0 8px 0;
        }
        
        .success-card p {
            color: #ffffff;
            font-size: 0.85rem;
            margin: 0;
        }
        
        /* Download section */
        .download-item {
            margin-bottom: 20px;
        }
        
        .download-label {
            font-size: 0.75rem;
            text-transform: uppercase;
            color: #ffffff;
            font-weight: 700;
            margin-bottom: 4px;
            letter-spacing: 1px;
        }
        
        .download-desc {
            font-size: 0.7rem;
            color: #d1d8e0;
            margin-bottom: 10px;
        }
        
        /* Footer */
        .app-footer {
            background: var(--slate-dark);
            border: 1px solid var(--border-subtle);
            border-radius: 4px;
            padding: 25px;
            text-align: center;
            margin-top: 60px;
        }
        
        .footer-text {
            font-size: 0.75rem;
            color: #ffffff;
            margin: 0 0 6px 0;
            letter-spacing: 0.5px;
        }
        
        .footer-credit {
            font-size: 0.7rem;
            color: #d1d8e0;
            margin: 0;
        }
        
        /* Spinner */
        .stSpinner > div {
            border-top-color: var(--accent-copper);
        }
        
        /* Hide Streamlit branding */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown("""
        <div class="app-header">
            <h1 class="app-title">Zoom Participation Tracker</h1>
            <p class="app-subtitle">Integrated Workspace</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Instructions
    with st.expander("📖 How to Use This Tool"):
        st.markdown("""
        **Follow these simple steps:**
        
        1. Enter your full name exactly as it appears in your Zoom display name
        2. Upload your Zoom transcript (`.txt` file)
        3. Upload your student list (`.xlsx` file with student names)
        4. Click 'Generate Reports' to create the PDFs
        5. Download your reports:
           - **Student Participation Report**: Share with students
           - **Teacher Analytics Report**: For your records
        
        *The Excel file name will be used as the course code in the report.*
        """)
    
    # Teacher name input
    st.markdown('<p class="section-label">Step 1: Enter Your Information</p>', unsafe_allow_html=True)
    
    teacher_name = st.text_input(
        "Your Full Name (as it appears in Zoom)",
        placeholder="e.g., MARIA JOSE GONZALEZ RODRIGUEZ",
        help="Enter your name exactly as it appears in your Zoom display name",
        key="teacher_name_input"
    )
    
    # File uploaders
    st.markdown('<p class="section-label">Step 2: Upload Files</p>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        transcript_file = st.file_uploader(
            "Zoom Transcript (.txt)",
            type=['txt'],
            help="Upload the Zoom meeting transcript file"
        )
    
    with col2:
        student_file = st.file_uploader(
            "Student List (.xlsx)",
            type=['xlsx'],
            help="Upload the Excel file with student names"
        )
    
    # Process button
    if teacher_name and transcript_file and student_file:
        st.markdown('<p class="section-label">Step 3: Generate Reports</p>', unsafe_allow_html=True)
        
        if st.button("Generate Reports"):
            try:
                with st.spinner("Processing attendance data..."):
                    # Extract course code from Excel filename
                    course_code = os.path.splitext(student_file.name)[0]
                    
                    # Parse transcript
                    st.info("📝 Analyzing Zoom transcript...")
                    word_counts, teacher_word_count = parse_transcript(transcript_file.read(), teacher_name)
                    transcript_file.seek(0)
                    
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
                    
                    # Calculate total student words
                    total_student_words = sum(word_count for _, word_count, _ in attendance_data)
                    
                    # Count statistics
                    present_count = sum(1 for _, _, status in attendance_data if status == "present")
                    absent_count = len(attendance_data) - present_count
                    
                    # Display summary
                    st.success("✅ Processing complete!")
                    
                    st.markdown('<p class="section-label">Summary Statistics</p>', unsafe_allow_html=True)
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Students", len(attendance_data))
                    col2.metric("Participated", present_count)
                    col3.metric("Absent", absent_count)
                    
                    # Generate Student Report PDF
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                        student_pdf_path = tmp_file.name
                        ecuador_tz = pytz.timezone('America/Guayaquil')
                        now_ecuador = datetime.now(ecuador_tz)
                        date_str = now_ecuador.strftime("%B %d, %Y")
                        create_student_report(attendance_data, student_pdf_path, date_str, course_code, teacher_name)
                    
                    with open(student_pdf_path, 'rb') as f:
                        student_pdf_data = f.read()
                    os.unlink(student_pdf_path)
                    
                    # Generate Teacher Analytics Report PDF
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                        teacher_pdf_path = tmp_file.name
                        create_teacher_analytics_report(teacher_word_count, total_student_words, teacher_pdf_path, date_str, course_code, teacher_name)
                    
                    with open(teacher_pdf_path, 'rb') as f:
                        teacher_pdf_data = f.read()
                    os.unlink(teacher_pdf_path)
                    
                    # Success message
                    st.markdown("""
                        <div class="success-card">
                            <h3>Reports Generated Successfully</h3>
                            <p>Your attendance reports are ready to download</p>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # Download buttons
                    st.markdown('<p class="section-label">Step 4: Download Your Reports</p>', unsafe_allow_html=True)
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown('<div class="download-item">', unsafe_allow_html=True)
                        st.markdown('<p class="download-label">Student Report</p>', unsafe_allow_html=True)
                        st.markdown('<p class="download-desc">Share with students</p>', unsafe_allow_html=True)
                        student_filename = f"Student_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                        st.download_button(
                            label="Download Student Report",
                            data=student_pdf_data,
                            file_name=student_filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown('<div class="download-item">', unsafe_allow_html=True)
                        st.markdown('<p class="download-label">Teacher Analytics</p>', unsafe_allow_html=True)
                        st.markdown('<p class="download-desc">For your records</p>', unsafe_allow_html=True)
                        teacher_filename = f"Teacher_Analytics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                        st.download_button(
                            label="Download Analytics",
                            data=teacher_pdf_data,
                            file_name=teacher_filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                        st.markdown('</div>', unsafe_allow_html=True)
                    
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.error("Please check your files and try again.")
    
    else:
        if not teacher_name:
            st.info("👆 Please enter your name to continue")
        elif not transcript_file or not student_file:
            st.info("👆 Please upload both files to continue")
    
    # Footer
    st.markdown("""
        <div class="app-footer">
            <p class="footer-text">Zoom Participation Tracker</p>
            <p class="footer-credit">Created by Jorge B. Cevallos Bravo</p>
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
