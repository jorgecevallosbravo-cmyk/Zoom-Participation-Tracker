# 📊 Zoom Participation Tracker

Automated participation tracking for Zoom classes. Analyzes transcripts to generate student participation reports and teacher analytics showing speaking contribution ratios.

**[🚀 Try the Live App](https://zoom-attendance.streamlit.app/)**

---
## How It Works

1. Upload your Zoom transcript (.txt)
2. Upload your student list (.xlsx)
3. Enter your name
4. Get two professional PDF reports:
   - **Student Report**: Who participated and word counts
   - **Teacher Analytics**: Teacher vs. student speaking contribution

---
## Input File Requirements
### Zoom Transcript (.txt)
✅ Standard Zoom transcript format (any language)

### Student List (.xlsx)
⚠️ **Important**: Excel file must have a column header containing **"ALUMNO"**

**Example:**

| ALUMNO                     | Other columns |
|----------------------------|---------------|
| GARCIA LOPEZ JUAN CARLOS   | ...           |
| PEREZ RODRIGUEZ MARIA JOSE | ...           |

**Why "ALUMNO"?**  
This tool was built for ESPAM MFL (Ecuador) using their institutional template format.

---
## Privacy
- No data storage
- Files processed in temporary memory only
- Everything deleted after you download reports
- No access logs

---
## Use Cases
- Track student engagement objectively
- Diagnose classroom balance (Are you talking too much?)
- Save 10+ hours per semester on manual tracking

---
## Built With
Python | Streamlit | ReportLab | Pandas

---
## Author
Jorge B. Cevallos  
Environmental Engineer & Automation Enthusiast  
*If it's repetitive, automate it.*
