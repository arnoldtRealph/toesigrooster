import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import base64
from collections import defaultdict
from datetime import datetime, timedelta
import uuid

# Updated admin periods data
data = {
    "Day 1": {
        "Period 1": ["M Cloete", "L Lubbe", "C Matthys", "J Saal", "PTL Smith", "G Zealand", "SSS"],
        "Period 2": ["P Gelderbloem", "J Klein", "PTL Smith", "W Van Wyk", "SSS"],
        "Period 3": ["K Ah Goo", "R Brandt", "D Van Eden", "R DE WAAL", "OC STEENKAMP", "SSS"],
        "Period 4": ["N Erasmus", "A Farmer", "T Lewis", "HL Jass", "C VAN WYK", "SSS"],
        "Period 5": ["N Farmer", "D Groenewald", "C Matthys", "PTL Smith", "R MOORE", "SSS"],
        "Period 6": ["J Clarke", "M Cloete", "A Isaacs", "PTL Smith", "A VAN ROOYEN", "SSS"],
        "Period 7": ["V Bovu", "S Erasmus", "D GROENEWALD", "HL Jass", "DT Van Wyk", "HL ISAKS", "SSS"]
    },
    "Day 2": {
        "Period 1": ["N Ful", "R Matthys", "K Witbooi", "G Zealand", "SSS"],
        "Period 2": ["J Klein", "N Dodds", "G Zealand", "R DE WAAL", "SSS"],
        "Period 3": ["E Erasmus", "A Farmer", "T Lewis", "PTL Smith", "OC STEENKAMP", "SSS"],
        "Period 4": ["N Dodds", "I TINTOA", "L Labbe", "HL Jass", "C VAN WYK", "SSS"],
        "Period 5": ["R Brandt", "J Clarke", "L Lubbe", "G Zealand", "OC STEENKAMP", "SSS"],
        "Period 6": ["V Bovu", "J Cloete", "T Lewis", "C Matthys", "A VAN ROOYEN", "SSS"],
        "Period 7": ["K Ah Goo", "A Farmer", "S Pier", "PTL Smith", "K Witbooi", "HL ISAKS", "SSS"]
    },
    "Day 3": {
        "Period 1": ["P GELDERBLOEM", "D GROENEWALD", "W Smith", "G Zealand", "SSS"],
        "Period 2": ["V Bovu", "N DODDS", "C Matthys", "PTL Smith", "SSS"],
        "Period 3": ["Y Goetzee", "E Erasmus", "J Klein", "M Matthys", "G Zealand", "SSS"],
        "Period 4": ["A Diedericks", "A Farmer", "S Pier", "A VAN ROOYEN", "PTL Smith", "C VAN WYK", "SSS"],
        "Period 5": ["P Gelderbloem", "J Klein", "C Matthys", "K Witbooi", "R MOORE", "SSS"],
        "Period 6": ["K Ah Goo", "M Cloete", "N DODDS", "C Matthys", "PTL KOOPMAN", "HL ISAKS", "SSS"],
        "Period 7": ["Y Coetzee", "W Lewis", "A Isaacs", "HL Jass", "DT Van Wyk", "SSS"]
    },
    "Day 4": {
        "Period 1": ["R Brandt", "T Lewis", "PTL Smith", "A ASSEGAAI", "SSS"],
        "Period 2": ["K Ah Goo", "E Cloete", "PTL Smith", "D Van Eden", "SSS"],
        "Period 3": ["K Ah Goo", "R Brandt", "D Van Eden", "B Rossouw", "OC STEENKAMP", "SSS"],
        "Period 4": ["V Bovu", "E Cloete", "A Farmer", "L Labbe", "C VAN WYK", "SSS"],
        "Period 5": ["Y Goetzee", "N DODDS", "T Lewis", "M Matthys", "A VAN ROOYEN", "SSS"],
        "Period 6": ["N Ful", "D GROENEWALD", "L Labbe", "PTL Smith", "R DE WAAL", "HL ISAKS", "SSS"],
        "Period 7": ["J Clarke", "A Diedericks", "A Isaacs", "S Pier", "PTL Smith", "A ASSEGAAI", "SSS"]
    },
    "Day 5": {
        "Period 1": ["E Cloete", "R Olivier", "G Zealand", "DT Van Wyk", "SSS"],
        "Period 2": ["Y Goetzee", "V Bovu", "J Klein", "PTL Smith", "SSS"],
        "Period 3": ["C Matthys", "PTL Smith", "D Van Eden", "P GELDERBLOEMK WITBOOI", "G Zealand", "SSS"],
        "Period 4": ["N DODDS", "A Isaacs", "R Olivier", "PTL Smith", "C VAN WYK", "SSS"],
        "Period 5": ["E Cloete", "PTL Smith", "I Stevens-Samuels", "W Van Wyk", "A ASSEGAAI", "SSS"],
        "Period 6": ["J Clarke", "M Erasmus", "A Farmer", "C Matthys", "HL Jass", "SSS"],
        "Period 7": ["Y Goetzee", "C Matthys", "M Matthys", "PTL Smith", "W Van Wyk", "SSS"]
    },
    "Day 6": {
        "Period 1": ["V Bovu", "A Farmer", "PTL Smith", "SSS"],
        "Period 2": ["Y Goetzee", "C Matthys", "PTL Smith", "A ASSEGAAI", "SSS"],
        "Period 3": ["R Brandt", "C Louderich", "A ASSEGAAI", "OC STEENKAMP", "H ISAKS", "SSS"],
        "Period 4": ["K Ab Goo", "E Erasmus", "S Pier", "D Van Eden", "C VAN WYK", "SSS"],
        "Period 5": ["J Clarke", "A Farmer", "P GELDERBLOEMK WITBOOI", "A KOOPMAN", "A VAN WYK", "SSS"],
        "Period 6": ["K Ab Goo", "M Cloete", "P GELDERBLOEM", "W Lewis", "PTL KOOPMAN", "SSS"],
        "Period 7": ["J Clarke", "D GROENEWALD", "S Erasmus", "HL Jass", "DT Van Wyk", "SSS"]
    }
}

# SMT teachers (including P GELDERBLOEM)
SMT_TEACHERS = [
    "AR VISAGIE", "C MATTHYS", "G ZEALAND", "J KLEIN",
    "Y GOETZEE", "I DIEDERICKS", "R BRANDT", "E CLOETE", "J SAAL",
    "ML MATTHYS", "M CLOETE", "P GELDERBLOEM", "D VAN EEDEN"
]

# Teachers to exclude entirely from substitutions
EXCLUDED_TEACHERS = ["PTL Smith"]

# Initialize session state for tracking absences, usage, and absence timestamps
if "absence_counts" not in st.session_state:
    st.session_state.absence_counts = defaultdict(int)
if "usage_counts" not in st.session_state:
    st.session_state.usage_counts = defaultdict(int)
if "daily_substitutes" not in st.session_state:
    st.session_state.daily_substitutes = set()
if "absence_timestamps" not in st.session_state:
    st.session_state.absence_timestamps = defaultdict(list)

# Streamlit app
st.title("TOESIGROOSTER PROGRAM")

# Get all unique educators
all_educators = set()
for day in data:
    for period in data[day]:
        all_educators.update(data[day][period])
all_educators = sorted(list(all_educators))

# Select absent educators, day, layout, and return periods
st.subheader("Absence and Return Management")
col1, col2, col3 = st.columns(3)
with col1:
    absent_educators = st.multiselect("Select Absent Educators", all_educators, help="Select one or more educators who are absent.")
with col2:
    selected_day = st.selectbox("Select Day", list(data.keys()))
with col3:
    day_layout = st.selectbox("Select Day Layout", ["2 Periods Before Eating/Break", "3 Periods Before Eating/Break"])

# Allow specifying return period for each absent educator
return_periods = {}
if absent_educators:
    st.subheader("Specify Return Periods for Absent Educators")
    for educator in absent_educators:
        return_period = st.selectbox(
            f"When does {educator} return?",
            ["Full Day Absence"] + [f"Period {i}" for i in range(1, 8)],
            key=f"return_{educator}_{uuid.uuid4()}"
        )
        return_periods[educator] = return_period

# Define the full schedule based on the selected layout
if day_layout == "2 Periods Before Eating/Break":
    full_schedule = ["Period 1", "Period 2", "Eating", "Break", "Period 3", "Period 4", "Period 5", "Break", "Period 6", "Period 7"]
    teaching_periods = ["Period 1", "Period 2", "Period 3", "Period 4", "Period 5", "Period 6", "Period 7"]
    period_mapping = {
        "Period 1": "Period 1",
        "Period 2": "Period 2",
        "Period 3": "Period 3",
        "Period 4": "Period 4",
        "Period 5": "Period 5",
        "Period 6": "Period 6",
        "Period 7": "Period 7"
    }
else:
    full_schedule = ["Period 1", "Period 2", "Period 3", "Eating", "Break", "Period 4", "Period 5", "Break", "Period 6", "Period 7"]
    teaching_periods = ["Period 1", "Period 2", "Period 3", "Period 4", "Period 5", "Period 6", "Period 7"]
    period_mapping = {
        "Period 1": "Period 1",
        "Period 2": "Period 2",
        "Period 3": "Period 3",
        "Period 4": "Period 4",
        "Period 5": "Period 5",
        "Period 6": "Period 6",
        "Period 7": "Period 7"
    }

# Function to select one substitute teacher, avoiding reuse if possible
def select_substitute(day, period, absent_educators, used_teachers):
    original_period = period_mapping[period]
    available_teachers = [t for t in data[day][original_period] if t not in absent_educators and t not in EXCLUDED_TEACHERS]
    if not available_teachers:
        return "No available teachers"
    
    # Split into non-SMT and SMT teachers
    non_smt_teachers = [t for t in available_teachers if t not in SMT_TEACHERS]
    smt_teachers = [t for t in available_teachers if t in SMT_TEACHERS]
    
    # First, try unused non-SMT teachers
    unused_non_smt = [t for t in non_smt_teachers if t not in used_teachers]
    if unused_non_smt:
        substitute = unused_non_smt[0]
    else:
        # Then, try reused non-SMT teachers
        if non_smt_teachers:
            substitute = non_smt_teachers[0]
        else:
            # Finally, try unused SMT teachers
            unused_smt = [t for t in smt_teachers if t not in used_teachers]
            if unused_smt:
                substitute = unused_smt[0]
            else:
                # Last resort: reused SMT teachers
                substitute = smt_teachers[0] if smt_teachers else "No available teachers"
    
    # Track usage for the selected substitute
    if substitute != "No available teachers":
        used_teachers.add(substitute)
        st.session_state.usage_counts[substitute] += 1
    return substitute

# Generate the substitute schedule document
if st.button("Generate Substitution Document"):
    # Reset daily substitutes for the new generation
    st.session_state.daily_substitutes = set()
    
    # Update absence counts and timestamps for absent educators based on return periods
    current_date = datetime.now()
    for educator in absent_educators:
        # Count absence only if they are absent for at least one period
        if return_periods.get(educator, "Full Day Absence") != f"Period 1":
            st.session_state.absence_counts[educator] += 1
            st.session_state.absence_timestamps[educator].append(current_date)
    
    # Create Word document
    doc = Document()
    
    # Set landscape orientation
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)  # A4 landscape width
    section.page_height = Inches(8.27)  # A4 landscape height
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    
    # Add title
    doc.add_heading("SAUL DAMON HIGH SCHOOL TOESIGROOSTER", 0)
    doc.add_paragraph(f"TOESIGROOSTER VIR {selected_day}")
    doc.add_paragraph(f"Day Layout: {day_layout}")
    
    # Create table: rows = 1 (header) + number of absent teachers (or 1 if none), columns = 1 (absent teachers) + full schedule
    num_rows = max(2, len(absent_educators) + 1)
    num_cols = 1 + len(full_schedule)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    
    # Function to set table borders
    def set_table_borders(table):
        tbl = table._element
        tbl_pr = tbl.tblPr
        tbl_borders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '8')  # Border thickness
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tbl_borders.append(border)
        tbl_pr.append(tbl_borders)
    
    # Apply borders to table
    set_table_borders(table)
    
    # Set header row with bold text
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Absent Teachers"
    for i, period in enumerate(full_schedule, 1):
        hdr_cells[i].text = period
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
    
    # Set column widths to fit landscape layout
    table.autofit = False
    for col in table.columns:
        for cell in col.cells:
            cell.width = Inches(1.0)  # Adjust width to fit 11 columns in landscape
    
    # Fill first column with absent teachers
    if absent_educators:
        for row_idx, teacher in enumerate(absent_educators, 1):
            table.rows[row_idx].cells[0].text = teacher
    else:
        table.rows[1].cells[0].text = "None"
    
    # Determine absent educators per period based on return periods
    period_order = [p for p in full_schedule if p in teaching_periods]
    for row_idx, teacher in enumerate(absent_educators, 1):
        return_period = return_periods.get(teacher, "Full Day Absence")
        if return_period == "Full Day Absence":
            periods_absent = period_order
        else:
            return_idx = int(return_period.split()[-1]) - 1
            periods_absent = period_order[:return_idx]
        
        for col_idx, period in enumerate(full_schedule, 1):
            if "Eating" in period or "Break" in period:
                table.rows[row_idx].cells[col_idx].text = ""
            elif period not in periods_absent:
                table.rows[row_idx].cells[col_idx].text = f"{teacher} (Returned)"
            else:
                # Only consider teachers absent for this period
                current_absent = [t for t in absent_educators if period in period_order and period_order.index(period) < (
                    int(return_periods[t].split()[-1]) if return_periods.get(t, "Full Day Absence") != "Full Day Absence" else len(period_order)
                )]
                substitute = select_substitute(selected_day, period, current_absent, st.session_state.daily_substitutes)
                table.rows[row_idx].cells[col_idx].text = substitute
    
    # Set font size for table cells (excluding header)
    for row_idx in range(1, num_rows):
        for cell in table.rows[row_idx].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    
    # Save document to bytes
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    # Download button with navy blue styling
    b64 = base64.b64encode(buffer.read()).decode()
    button_style = """
        <style>
            .download-btn {
                display: inline-block;
                padding: 10px 20px;
                background-color: #1C2526;
                color: white;
                text-decoration: none;
                border-radius: 5px;
                font-weight: bold;
                text-align: center;
            }
            .download-btn:hover {
                background-color: #2A3D45;
            }
        </style>
    """
    href = f'{button_style}<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="substitute_schedule.docx" class="download-btn">Download Substitute Schedule</a>'
    st.markdown(href, unsafe_allow_html=True)

# Insights and Visualizations
st.header("INSIGTE EN VISUALISERING")

# Absence Frequency Graph
st.subheader("OPVOEDER AFWESIGHEIDS FREKWENSIE (OVERALL)")
absence_data = pd.Series(st.session_state.absence_counts)
if not absence_data.empty:
    fig, ax = plt.subplots()
    absence_data.plot(kind='bar', ax=ax)
    plt.title("Number of Absences per Educator")
    plt.xlabel("Educator")
    plt.ylabel("Number of Absences")
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig)
else:
    st.write("No absence data available yet.")

# Absenteeism Summary
st.subheader("AFWESIGHEIDSOPSOMMING")
summary_period = st.selectbox("Select Summary Period", ["Weekly", "Monthly", "Quarterly"])

# Calculate the date range for the summary
current_date = datetime.now()
if summary_period == "Weekly":
    start_date = current_date - timedelta(days=7)
    period_label = "Last 7 Days"
elif summary_period == "Monthly":
    start_date = current_date - timedelta(days=30)
    period_label = "Last 30 Days"
else:  # Quarterly
    start_date = current_date - timedelta(days=90)
    period_label = "Last 90 Days"

# Filter absences within the selected period
period_absences = defaultdict(int)
for educator, timestamps in st.session_state.absence_timestamps.items():
    for ts in timestamps:
        if ts >= start_date:
            period_absences[educator] += 1

# Display the summary as a bar chart
st.subheader(f"OPVOEDER AFWESIGHEIDS FREKWENSIE ({period_label})")
period_absence_data = pd.Series(period_absences)
if not period_absence_data.empty:
    fig, ax = plt.subplots()
    period_absence_data.plot(kind='bar', ax=ax)
    plt.title(f"Number of Absences per Educator ({period_label})")
    plt.xlabel("Educator")
    plt.ylabel("Number of Absences")
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig)
else:
    st.write(f"No absence data available for the {period_label.lower()}.")

# Substitute Usage Table
st.subheader("Substitute Usage Summary")
usage_data = pd.Series(st.session_state.usage_counts)
if not usage_data.empty:
    # Create a DataFrame for the table
    usage_df = pd.DataFrame(usage_data.items(), columns=["Educator", "Number of Substitutions"])
    usage_df = usage_df.sort_values(by="Number of Substitutions", ascending=False)
    usage_df = usage_df.reset_index(drop=True)
    # Style the table
    st.table(usage_df.style.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#4CAF50'), ('color', 'white'), ('font-weight', 'bold')]},
        {'selector': 'td', 'props': [('border', '1px solid #ddd'), ('padding', '8px')]},
        {'selector': 'tr:nth-child(even)', 'props': [('background-color', '#f2f2f2')]}
    ]))
else:
    st.write("No substitution data available yet.")