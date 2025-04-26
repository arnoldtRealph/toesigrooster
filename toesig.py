import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import base64
from collections import defaultdict
from datetime import datetime, timedelta
import uuid
import logging

# Set wide layout for the entire app
st.set_page_config(layout="wide")

# Set up logging for debugging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Load the admin periods CSV file from the root folder
df = pd.read_csv("admin periods.csv")

# Convert CSV to dictionary format
data = {}
for day in df.columns[1:]:
    data[day] = {}
    for _, row in df.iterrows():
        period = row['Period']
        teachers = row[day]
        # Clean and split teacher names, remove empty strings
        teacher_list = [t.strip() for t in teachers.split(',') if t.strip()]
        data[day][period] = teacher_list

# Get all unique educators from CSV for dropdown menu
all_educators = set()
for day in data:
    for period in data[day]:
        all_educators.update(data[day][period])
all_educators = sorted(list(all_educators))
logger.debug(f"All educators: {all_educators}")

# SMT teachers
SMT_TEACHERS = [
    "AR VISAGIE", "C MATTHYS", "G ZEALAND", "J KLEIN",
    "Y COETZEE", "I DIEDERICKS", "R BRANDT", "E CLOETE", "J SAAL",
    "ML MATTHYS", "M CLOETE", "P GELDERBLOEM", "D VAN EEDEN"
]

# Teachers to exclude from substitutions
EXCLUDED_TEACHERS = []

# Initialize session state
if "absence_counts" not in st.session_state:
    st.session_state.absence_counts = defaultdict(int)
if "usage_counts" not in st.session_state:
    st.session_state.usage_counts = defaultdict(int)
if "usage_timestamps" not in st.session_state:
    st.session_state.usage_timestamps = defaultdict(list)
if "daily_substitutes" not in st.session_state:
    st.session_state.daily_substitutes = defaultdict(set)
if "absence_timestamps" not in st.session_state:
    st.session_state.absence_timestamps = defaultdict(list)

# Color scheme: Blues, whites, and green accents
st.markdown("""
    <style>
        .main {
            background-color: #F5F7FA;
            padding: 20px;
            border-radius: 10px;
            max-width: 95vw !important;
            width: 100%;
            margin: 0 auto;
        }
        .stButton>button {
            background-color: #005B99;
            color: white;
            border-radius: 5px;
            padding: 10px 20px;
            font-weight: bold;
        }
        .stButton>button:hover {
            background-color: #007ACC;
        }
        .generate-button>button, .download-button>button {
            background-color: #28A745 !important;
            color: white !important;
            border-radius: 5px !important;
            padding: 10px 20px !important;
            font-weight: bold !important;
            border: none !important;
        }
        .generate-button>button:hover, .download-button>button:hover {
            background-color: #218838 !important;
        }
        h1, h2, h3 {
            color: #003087;
            text-align: center;
        }
        .section {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            max-width: 95vw !important;
            width: 100%;
            margin: 0 auto;
        }
        .table-container {
            overflow-x: auto;
            width: 100%;
        }
        .wide-table {
            width: 100%;
            max-width: 1600px !important;
            margin: 0 auto;
            font-size: 18px;
            border-collapse: collapse;
            word-spacing: 3px;
            white-space: nowrap;
            background-color: #FFFFFF;
        }
        .wide-table th, .wide-table td {
            padding: 15px !important;
            text-align: center;
            border: 1px solid #005B99;
        }
        .wide-table th {
            background-color: #005B99;
            color: white;
            font-weight: bold;
            font-size: 20px;
        }
        .wide-table td {
            background-color: #E6F0FA;
            color: #003087;
        }
        .wide-table tr:nth-child(even) td {
            background-color: #D1E3F5;
        }
        .wide-table tr:hover td {
            background-color: #B3D4FC;
            transition: background-color 0.2s;
        }
        .stSelectbox, .stMultiSelect, .stNumberInput {
            width: 90% !important;
            max-width: none !important;
        }
        .stColumn > div {
            display: flex;
            justify-content: center;
        }
    </style>
""", unsafe_allow_html=True)

# Streamlit app
st.title("SAUL DAMON HIGH SCHOOL TOESIGROOSTER")
st.markdown("<div class='main'>", unsafe_allow_html=True)

# Input Section
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.subheader("Afwesigheid en Skedule Konfigurasie")
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    absent_educators = st.multiselect("Kies Afwesige Opvoeders", all_educators, help="Kies opvoeders wat afwesig is.")
with col2:
    selected_day = st.selectbox("Kies Dag", list(data.keys()))
with col3:
    day_layout = st.selectbox("Kies Dag Uitleg", ["2 Periodes Voor Eet/Break", "3 Periodes Voor Eet/Break"])
with col4:
    start_period = st.selectbox("Begin Periode", [f"Periode {i}" for i in range(1, 8)], index=0)
with col5:
    num_periods = st.number_input("Aantal Periodes", min_value=1, max_value=7, value=7, step=1)

# Clear all inputs button
if st.button("Maak Alle Insette Skoon"):
    st.session_state.clear()
    st.rerun()

# Define current_day
current_day = selected_day

# Return periods for absent educators
return_periods = {}
if absent_educators:
    st.subheader("Spesifiseer Terugkeer Periodes")
    start_period_idx = int(start_period.split()[-1])
    for educator in absent_educators:
        return_period = st.selectbox(
            f"Wanneer keer {educator} terug?",
            ["Volle Dag Afwesig"] + [f"Periode {i}" for i in range(start_period_idx, 8)],
            key=f"return_{educator}_{uuid.uuid4()}"
        )
        return_periods[educator] = return_period
st.markdown("</div>", unsafe_allow_html=True)

# Generate dynamic schedule based on start period and number of periods
start_period_idx = int(start_period.split()[-1])
period_mapping = {f"Periode {i}": f"Period {i}" for i in range(1, 8)}
days = list(data.keys())
current_day_idx = days.index(selected_day)
next_day_idx = (current_day_idx + 1) % len(days)
next_day = days[next_day_idx]

# Calculate teaching periods and full schedule
teaching_periods = []
full_schedule = []
current_period_idx = start_period_idx
periods_added = 0
teaching_periods_count = 0  # Count teaching periods for break insertion
teaching_periods_since_pouse1 = 0  # Count teaching periods since POUSE 1

# Determine how many periods before "Eet" based on day_layout
periods_before_eat = 2 if day_layout == "2 Periodes Voor Eet/Break" else 3
# Determine how many periods between POUSE 1 and POUSE 2
periods_before_pouse2 = 3 if day_layout == "2 Periodes Voor Eet/Break" else 2

while periods_added < num_periods:
    period_num = ((current_period_idx - 1) % 7) + 1 if current_period_idx > 7 else current_period_idx
    period_name = f"Periode {period_num}"
    if current_period_idx > 7:
        period_name += " (Dag 2)"
    teaching_periods.append(period_name)
    
    # Add the period to the full schedule
    full_schedule.append(period_name)
    teaching_periods_count += 1
    teaching_periods_since_pouse1 += 1
    periods_added += 1
    current_period_idx += 1
    
    # Insert "Eet" and "POUSE 1" after the specified number of teaching periods
    if teaching_periods_count == periods_before_eat and periods_added < num_periods:
        full_schedule.extend(["Eet", "POUSE 1"])
        teaching_periods_since_pouse1 = 0
    
    # Insert "POUSE 2" after the specified number of teaching periods since POUSE 1
    if teaching_periods_since_pouse1 == periods_before_pouse2 and periods_added < num_periods:
        full_schedule.append("POUSE 2")
        teaching_periods_since_pouse1 = 0

# Function to select substitute teacher
def select_substitute(selected_day, period, absent_educators, used_teachers, start_period_idx):
    try:
        period_num = int(period.split()[1].split(" ")[0])
        # Determine if the period belongs to the next day
        period_position = teaching_periods.index(period) + 1  # 1-based index in the schedule
        total_periods_first_day = 7 - start_period_idx + 1
        day_to_use = selected_day if period_position <= total_periods_first_day else next_day
        original_period = f"Period {period_num}"
        scheduled_teachers = data.get(day_to_use, {}).get(original_period, [])
        logger.debug(f"Day: {day_to_use}, Period: {period}, Original Period: {original_period}")
        logger.debug(f"Scheduled teachers for {day_to_use}, {original_period}: {scheduled_teachers}")
        logger.debug(f"Absent educators: {absent_educators}")
        logger.debug(f"Excluded teachers: {EXCLUDED_TEACHERS}")
        logger.debug(f"Used teachers for {period}: {used_teachers}")

        # Filter available teachers
        available_teachers = [
            t for t in scheduled_teachers
            if t not in absent_educators and t not in EXCLUDED_TEACHERS and t not in used_teachers
        ]
        logger.debug(f"Available teachers after filtering: {available_teachers}")

        if not available_teachers:
            logger.warning(f"No available teachers for {day_to_use}")
            return "OPDEEL"
        
        # Prioritize non-SMT teachers, then SMT teachers
        non_smt_teachers = [t for t in available_teachers if t not in SMT_TEACHERS]
        smt_teachers = [t for t in available_teachers if t in SMT_TEACHERS]
        
        if non_smt_teachers:
            substitute = non_smt_teachers[0]
            logger.debug(f"Selected non-SMT substitute: {substitute}")
        elif smt_teachers:
            substitute = smt_teachers[0]
            logger.debug(f"Selected SMT substitute: {substitute}")
        else:
            substitute = available_teachers[0]
            logger.debug(f"Selected general substitute: {substitute}")
        
        if substitute != "OPDEEL":
            used_teachers.add(substitute)
            st.session_state.usage_counts[substitute] += 1
            st.session_state.usage_timestamps[substitute].append((datetime.now(), day_to_use))
            logger.debug(f"Updated used_teachers for {period}: {used_teachers}")
        return substitute
    except KeyError as e:
        logger.error(f"KeyError in select_substitute: {str(e)}")
        return "OPDEEL"
    except Exception as e:
        logger.error(f"Error in select_substitute: {str(e)}")
        return "OPDEEL"

# Create unique column names
unique_columns = ["Afwesige Opvoeders"] + full_schedule

# Generate substitution schedule once
st.session_state.daily_substitutes = defaultdict(set)  # Reset substitutes
num_rows = max(2, len(absent_educators) + 1)
num_cols = len(unique_columns)
table_data = [["" for _ in range(num_cols)] for _ in range(num_rows)]
table_data[0] = unique_columns

if absent_educators:
    for row_idx, teacher in enumerate(absent_educators, 1):
        table_data[row_idx][0] = teacher
else:
    table_data[1][0] = "Geen"

period_order = [p for p in full_schedule if p in teaching_periods]
for row_idx, teacher in enumerate(absent_educators, 1):
    return_period = return_periods.get(teacher, "Volle Dag Afwesig")
    if return_period == "Volle Dag Afwesig":
        periods_absent = period_order
    else:
        return_idx = int(return_period.split()[-1])
        periods_absent = [p for p in period_order if int(p.split()[1].split(" ")[0]) < return_idx]
    
    for col_idx, period in enumerate(full_schedule, 1):
        if period not in teaching_periods:
            table_data[row_idx][col_idx] = ""
        elif period not in periods_absent:
            table_data[row_idx][col_idx] = f"{teacher} (Terug)"
        else:
            current_absent = [
                t for t in absent_educators
                if period in period_order and (
                    return_periods.get(t, "Volle Dag Afwesig") == "Volle Dag Afwesig" or
                    int(return_periods[t].split()[-1]) > int(period.split()[1].split(" ")[0])
                )
            ]
            substitute = select_substitute(selected_day, period, current_absent, st.session_state.daily_substitutes[period], start_period_idx)
            table_data[row_idx][col_idx] = substitute

# Substitution Schedule Table
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.subheader("TOESIGROOSTER")
df_schedule = pd.DataFrame(table_data[1:], columns=unique_columns)
df_schedule.index = df_schedule.index + 1  # Start index from 1

# Ensure columns are unique by appending a suffix if necessary
unique_cols = []
seen = {}
for col in unique_columns:
    if col in seen:
        seen[col] += 1
        unique_cols.append(f"{col}_{seen[col]}")
    else:
        seen[col] = 0
        unique_cols.append(col)
df_schedule.columns = unique_cols

# Apply styling
st.markdown("<div class='table-container'><div class='wide-table'>", unsafe_allow_html=True)
st.table(df_schedule.style.set_properties(**{
    'background-color': '#E6F0FA',
    'color': '#003087',
    'border': '1px solid #005B99',
    'padding': '15px',
    'text-align': 'center',
    'font-size': '18px',
    'word-spacing': '3px',
    'white-space': 'nowrap'
}).set_table_styles([
    {'selector': 'th', 'props': [('background-color', '#005B99'), ('color', 'white'), ('font-weight', 'bold'), ('font-size', '20px'), ('padding', '15px')]},
    {'selector': 'tr:nth-child(even)', 'props': [('background-color', '#D1E3F5')]},
    {'selector': 'tr:hover', 'props': [('background-color', '#B3D4FC')]}
]))
st.markdown("</div></div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# Available Teachers Table
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.subheader("Beskikbare Opvoeders per Periode")
available_data = {}
for period in teaching_periods:
    period_num = int(period.split()[1].split(" ")[0])
    period_position = teaching_periods.index(period) + 1
    total_periods_first_day = 7 - start_period_idx + 1
    day_to_use = current_day if period_position <= total_periods_first_day else next_day
    available_data[period] = ", ".join([
        t for t in data[day_to_use][f"Period {period_num}"]
        if t not in absent_educators and t not in EXCLUDED_TEACHERS and t not in st.session_state.daily_substitutes[period]
    ] or ["Geen"])
df_available = pd.DataFrame(list(available_data.items()), columns=["Periode", "Beskikbare Opvoeders"])
df_available.index = df_available.index + 1  # Start index from 1
st.table(df_available.style.set_properties(**{
    'background-color': '#E6F0FA',
    'color': '#003087',
    'border': '1px solid #005B99',
    'padding': '10px',
    'text-align': 'center',
    'font-size': '16px'
}).set_table_styles([{
    'selector': 'th',
    'props': [('background-color', '#005B99'), ('color', 'white'), ('font-weight', 'bold'), ('font-size', '16px'), ('padding', '10px')]
}]))
st.markdown("</div>", unsafe_allow_html=True)

# Generate the substitute schedule document
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.markdown("<div class='generate-button'>", unsafe_allow_html=True)
if st.button("Genereer TOESIGROOSTER"):
    current_date = datetime.now()
    for educator in absent_educators:
        if return_periods.get(educator, "Volle Dag Afwesig") != f"Periode {start_period_idx}":
            st.session_state.absence_counts[educator] += 1
            st.session_state.absence_timestamps[educator].append(current_date)
    
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    
    # Header
    doc.add_heading("SAUL DAMON HIGH SCHOOL", 0).alignment = 1
    p = doc.add_paragraph()
    run = p.add_run(f"TOESIGROOSTER VIR {selected_day} ({datetime.now().strftime('%d/%m/%Y')})")
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 48, 135)
    p.alignment = 1
    doc.add_paragraph(f"Dag Uitleg: {day_layout} | Begin by: {start_period} | Aantal Periodes: {num_periods}").alignment = 1
    
    # Substitution Table
    num_rows = len(absent_educators) + 1 if absent_educators else 2
    num_cols = len(unique_columns)
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    table.autofit = False
    
    def set_table_borders(table):
        tbl = table._element
        tbl_pr = tbl.tblPr
        tbl_borders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:color'), '005B99')
            tbl_borders.append(border)
        tbl_pr.append(tbl_borders)
    
    set_table_borders(table)
    
    # Set header row
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(unique_columns):
        hdr_cells[i].text = header
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(255, 255, 255)
            paragraph.paragraph_format.space_after = Pt(0)
        cell_fill = OxmlElement('w:shd')
        cell_fill.set(qn('w:fill'), '005B99')
        hdr_cells[i]._element.get_or_add_tcPr().append(cell_fill)
    
    # Calculate column width to fit page
    page_width = section.page_width - section.left_margin - section.right_margin
    col_width = page_width / num_cols
    for col in table.columns:
        for cell in col.cells:
            cell.width = col_width
    
    # Populate data rows from table_data
    if absent_educators:
        for row_idx, teacher in enumerate(absent_educators, 1):
            for col_idx in range(num_cols):
                cell = table.rows[row_idx].cells[col_idx]
                cell.text = table_data[row_idx][col_idx]
                if row_idx % 2 == 0:
                    cell_fill = OxmlElement('w:shd')
                    cell_fill.set(qn('w:fill'), 'E6F0FA')
                    cell._element.get_or_add_tcPr().append(cell_fill)
    else:
        table.rows[1].cells[0].text = "Geen"
    
    for row_idx in range(1, num_rows):
        for cell in table.rows[row_idx].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(0, 48, 135)
    
    # Available Teachers Section
    doc.add_heading("Beskikbare Opvoeders per Periode", level=2)
    for period in teaching_periods:
        period_num = int(period.split()[1].split(" ")[0])
        period_position = teaching_periods.index(period) + 1
        total_periods_first_day = 7 - start_period_idx + 1
        day_to_use = current_day if period_position <= total_periods_first_day else next_day
        available = [
            t for t in data[day_to_use][f"Period {period_num}"]
            if t not in absent_educators and t not in EXCLUDED_TEACHERS and t not in st.session_state.daily_substitutes[period]
        ]
        doc.add_paragraph(f"{period}: {', '.join(available) if available else 'Geen'}")
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.markdown("<div class='download-button'>", unsafe_allow_html=True)
    st.download_button(
        label="Download TOESIGROOSTER",
        data=buffer,
        file_name="vervangingskedule.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# Insights Section
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.header("INSIGTE EN VISUALISERING")

# Absence Frequency Graph
st.subheader("OPVOEDER AFWESIGHEIDS FREKWENSIE (ALGEHEEL)")
absence_data = pd.Series(st.session_state.absence_counts)
if not absence_data.empty:
    fig, ax = plt.subplots(figsize=(12, 6))
    absence_data.plot(kind='bar', ax=ax, color='#005B99')
    plt.title("Aantal Afwesighede per Opvoeder", fontsize=14)
    plt.xlabel("Opvoeder", fontsize=12)
    plt.ylabel("Aantal Afwesighede", fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig)
else:
    st.write("Geen afwesigheidsdata beskikbaar nie.")

# Absenteeism Summary
st.subheader("AFWESIGHEIDSOPSOMMING")
summary_period = st.selectbox("Kies Opsommingsperiode", ["Weekliks", "Maandelikse", "Kwartaalliks"])
current_date = datetime.now()
if summary_period == "Weekliks":
    start_date = current_date - timedelta(days=7)
    period_label = "Laaste 7 Dae"
elif summary_period == "Maandelikse":
    start_date = current_date - timedelta(days=30)
    period_label = "Laaste 30 Dae"
else:
    start_date = current_date - timedelta(days=90)
    period_label = "Laaste 90 Dae"

period_absences = defaultdict(int)
for educator, timestamps in st.session_state.absence_timestamps.items():
    for ts in timestamps:
        if ts >= start_date:
            period_absences[educator] += 1

st.subheader(f"OPVOEDER AFWESIGHEIDS FREKWENSIE ({period_label})")
period_absence_data = pd.Series(period_absences)
if not period_absence_data.empty:
    fig, ax = plt.subplots(figsize=(12, 6))
    period_absence_data.plot(kind='bar', ax=ax, color='#005B99')
    plt.title(f"Aantal Afwesighede per Opvoeder ({period_label})", fontsize=14)
    plt.xlabel("Opvoeder", fontsize=12)
    plt.ylabel("Aantal Afwesighede", fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig)
else:
    st.write(f"Geen afwesigheidsdata vir die {period_label.lower()} beskikbaar nie.")

# Substitute Usage Table
st.markdown("<div class='section'>", unsafe_allow_html=True)
st.subheader(f"Opvoeder Frekwensie ({period_label})")
period_usage = defaultdict(int)
for educator, timestamps in st.session_state.usage_timestamps.items():
    for ts, _ in timestamps:
        if ts >= start_date:
            period_usage[educator] += 1

usage_data = pd.Series(period_usage)
if not usage_data.empty:
    usage_df = pd.DataFrame(usage_data.items(), columns=["Opvoeder", "Aantal Vervangings"])
    usage_df = usage_df.sort_values(by="Aantal Vervangings", ascending=False).reset_index(drop=True)
    usage_df.index = usage_df.index + 1  # Start index from 1
    st.table(usage_df.style.set_properties(**{
        'background-color': '#E6F0FA',
        'color': '#003087',
        'border': '1px solid #005B99',
        'padding': '10px',
        'text-align': 'center',
        'font-size': '16px'
    }).set_table_styles([{
        'selector': 'th',
        'props': [('background-color', '#005B99'), ('color', 'white'), ('font-weight', 'bold'), ('font-size', '16px'), ('padding', '10px')]
    }]))
else:
    st.write(f"Geen vervangingsdata vir die {period_label.lower()} beskikbaar nie.")
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)