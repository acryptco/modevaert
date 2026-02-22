import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
import datetime

def parse_members(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)
    members = df.iloc[2:, 0].dropna().tolist()
    return members

def normalize_name(name):
    """Normalize a name for comparison by removing extra spaces and converting to lowercase"""
    return ' '.join(name.split()).lower()

def find_matching_member(pdf_name, members_list):
    """Find a matching member name from the members list, handling variations"""
    normalized_pdf_name = normalize_name(pdf_name)
    
    # First try exact match
    for member in members_list:
        if normalize_name(member) == normalized_pdf_name:
            return member
    
    # Try partial matches (e.g., "Michael Vollenberg Keler" matches "Michael Keler")
    for member in members_list:
        normalized_member = normalize_name(member)
        # Check if the member name is contained in the PDF name or vice versa
        if (normalized_member in normalized_pdf_name or 
            normalized_pdf_name in normalized_member):
            return member
    
    # Try matching by last name (most reliable for Danish names)
    pdf_words = normalized_pdf_name.split()
    for member in members_list:
        normalized_member = normalize_name(member)
        member_words = normalized_member.split()
        
        # Check if any word from PDF name matches any word from member name
        for pdf_word in pdf_words:
            for member_word in member_words:
                if pdf_word == member_word and len(pdf_word) > 2:  # Avoid matching short words
                    return member
    
    return None

def calculate_weekday_in_range(start_day, end_day, month_name, year, target_weekday):
    """
    Calculate the date of a specific weekday within a date range.
    target_weekday: 0=Monday, 1=Tuesday, ..., 6=Sunday
    """
    month_order = {
        'Januar': 1, 'Februar': 2, 'Marts': 3, 'April': 4, 'Maj': 5, 'Juni': 6,
        'Juli': 7, 'August': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'December': 12
    }
    
    month_num = month_order.get(month_name, 1)
    
    # Find the target weekday in the range
    for day in range(start_day, end_day + 1):
        try:
            date_obj = datetime.date(year, month_num, day)
            if date_obj.weekday() == target_weekday:
                return date_obj
        except ValueError:
            continue
    return None

def parse_program(uploaded_files, members_list):
    all_meetings = {}
    
    for uploaded_file in uploaded_files:
        with pdfplumber.open(uploaded_file) as pdf:
            text = '\n'.join(page.extract_text() for page in pdf.pages if page.extract_text())
        meetings = {}
        current_date = None
        current_weekend_date = None
        assigned = set()
        weekend_assigned = set()
        in_weekend_section = False  # Track if we're in a weekend meeting section
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            
            # Look for weekday meeting dates (e.g., "Tirsdag 15 September", "Torsdag 09 December")
            weekday_date_match = re.match(r'(Tirsdag|Mandag|Torsdag) \d{2} (September|Oktober|November|December|Januar|Februar|Marts|April|Maj|Juni|Juli|August)', line)
            
            # Look for weekend meeting dates (e.g., "07/09/2025")
            weekend_date_match = re.match(r'(\d{2})/(\d{2})/(\d{4})', line)
            
            # Look for January format dates (e.g., "06. JAN | UGENS BIBELLÆSNING: ESAJAS 17-20")
            january_date_match = re.search(r'(\d{2})\.\s*JAN', line)
            
            # Look for other Danish date formats (e.g., "01 Januar 26", "01. Januar 26")
            danish_date_match = re.search(r'(\d{2})\.?\s*(Januar|Februar|Marts|April|Maj|Juni|Juli|August|September|Oktober|November|December)', line)

            # Look for abbreviated Danish month formats (e.g., "03. FEB", "10 FEB")
            abbrev_date_match = re.search(
                r'(\d{2})\.?\s*(JAN|FEB|MAR|APR|MAJ|JUN|JUL|AUG|SEP|OKT|NOV|DEC)',
                line,
                re.IGNORECASE
            )
            
            # NEW: Look for date range format (e.g., "marts 02-08")
            date_range_match = re.search(
                r'(januar|februar|marts|april|maj|juni|juli|august|september|oktober|november|december)\s+(\d{1,2})-(\d{1,2})',
                line,
                re.IGNORECASE
            )
            
            # NEW: Look for cross-month date range (e.g., "Marts 30-april 05")
            cross_month_range_match = re.search(
                r'(januar|februar|marts|april|maj|juni|juli|august|september|oktober|november|december)\s+(\d{1,2})-(januar|februar|marts|april|maj|juni|juli|august|september|oktober|november|december)\s+(\d{1,2})',
                line,
                re.IGNORECASE
            )
            
            if cross_month_range_match:
                # Handle cross-month ranges first (more specific)
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                
                # Reset weekend section flag when we encounter a new date range
                in_weekend_section = False
                
                start_month_str, start_day_str, end_month_str, end_day_str = cross_month_range_match.groups()
                
                # Capitalize month names
                month_capitalize = {
                    'januar': 'Januar', 'februar': 'Februar', 'marts': 'Marts', 'april': 'April',
                    'maj': 'Maj', 'juni': 'Juni', 'juli': 'Juli', 'august': 'August',
                    'september': 'September', 'oktober': 'Oktober', 'november': 'November', 'december': 'December'
                }
                start_month = month_capitalize.get(start_month_str.lower(), start_month_str)
                end_month = month_capitalize.get(end_month_str.lower(), end_month_str)
                
                year = 2026
                start_day = int(start_day_str)
                end_day = int(end_day_str)
                
                # Calculate Tuesday in the range (1 = Tuesday)
                tuesday_date = calculate_weekday_in_range(start_day, 31, start_month, year, 1)
                if not tuesday_date:
                    tuesday_date = calculate_weekday_in_range(1, end_day, end_month, year, 1)
                
                # Calculate Sunday in the range (6 = Sunday)
                sunday_date = calculate_weekday_in_range(start_day, 31, start_month, year, 6)
                if not sunday_date:
                    sunday_date = calculate_weekday_in_range(1, end_day, end_month, year, 6)
                
                current_date = None
                current_weekend_date = None
                
                if tuesday_date:
                    month_order = {
                        1: 'Januar', 2: 'Februar', 3: 'Marts', 4: 'April', 5: 'Maj', 6: 'Juni',
                        7: 'Juli', 8: 'August', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'December'
                    }
                    month_name = month_order[tuesday_date.month]
                    current_date = f"Tirsdag {tuesday_date.day:02d} {month_name} {tuesday_date.year}"
                    assigned = set()
                
                if sunday_date:
                    month_order = {
                        1: 'Januar', 2: 'Februar', 3: 'Marts', 4: 'April', 5: 'Maj', 6: 'Juni',
                        7: 'Juli', 8: 'August', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'December'
                    }
                    month_name = month_order[sunday_date.month]
                    current_weekend_date = f"Søndag {sunday_date.day:02d} {month_name} {sunday_date.year}"
                    weekend_assigned = set()
                
                if 'Intet møde' in line or 'Ingen møde' in line:
                    current_date = None
                    current_weekend_date = None
                    continue
            elif date_range_match:
                # Handle single-month date ranges
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                
                # Reset weekend section flag when we encounter a new date range
                in_weekend_section = False
                
                month_str, start_day_str, end_day_str = date_range_match.groups()
                
                # Capitalize month names
                month_capitalize = {
                    'januar': 'Januar', 'februar': 'Februar', 'marts': 'Marts', 'april': 'April',
                    'maj': 'Maj', 'juni': 'Juni', 'juli': 'Juli', 'august': 'August',
                    'september': 'September', 'oktober': 'Oktober', 'november': 'November', 'december': 'December'
                }
                month_name = month_capitalize.get(month_str.lower(), month_str)
                
                year = 2026
                start_day = int(start_day_str)
                end_day = int(end_day_str)
                
                # Calculate Tuesday (weekly meeting) - weekday 1
                tuesday_date = calculate_weekday_in_range(start_day, end_day, month_name, year, 1)
                
                # Calculate Sunday (weekend meeting) - weekday 6
                sunday_date = calculate_weekday_in_range(start_day, end_day, month_name, year, 6)
                
                current_date = None
                current_weekend_date = None
                
                if tuesday_date:
                    current_date = f"Tirsdag {tuesday_date.day:02d} {month_name} {tuesday_date.year}"
                    assigned = set()
                
                if sunday_date:
                    current_weekend_date = f"Søndag {sunday_date.day:02d} {month_name} {sunday_date.year}"
                    weekend_assigned = set()
                
                if 'Intet møde' in line or 'Ingen møde' in line:
                    current_date = None
                    current_weekend_date = None
                    continue
            elif weekday_date_match:
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                # Reset weekend section flag when we encounter a new date
                in_weekend_section = False
                # Add year to weekday dates for consistency
                weekday_date = weekday_date_match.group(0)
                current_date = f"{weekday_date} 2025"
                assigned = set()
                weekend_assigned = set()
                current_weekend_date = None
                if 'Ingen møde' in line:
                    current_date = None
                    continue
            elif weekend_date_match:
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                # Convert DD/MM/YYYY to a readable format
                day, month, year = weekend_date_match.groups()
                month_names = {
                    '01': 'Januar', '02': 'Februar', '03': 'Marts', '04': 'April',
                    '05': 'Maj', '06': 'Juni', '07': 'Juli', '08': 'August',
                    '09': 'September', '10': 'Oktober', '11': 'November', '12': 'December'
                }
                month_name = month_names.get(month, month)
                current_date = f"Søndag {day} {month_name} {year}"
                assigned = set()
                weekend_assigned = set()
                current_weekend_date = None
            elif january_date_match:
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                # Handle January format (e.g., "06. JAN")
                day = january_date_match.group(1)
                current_date = f"Tirsdag {day} Januar 2026"
                assigned = set()
                weekend_assigned = set()
                current_weekend_date = None
                if 'Ingen møde' in line:
                    current_date = None
                    continue
            elif danish_date_match:
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                # Handle other Danish date formats (e.g., "01 Januar 26")
                day, month = danish_date_match.groups()
                # Determine year - assume 2026 for January dates, 2025 for others
                year = "2026" if month == "Januar" else "2025"
                current_date = f"Tirsdag {day} {month} {year}"
                assigned = set()
                weekend_assigned = set()
                current_weekend_date = None
                if 'Ingen møde' in line:
                    current_date = None
                    continue
            elif abbrev_date_match:
                if current_date:
                    meetings[current_date] = assigned
                if current_weekend_date:
                    meetings[current_weekend_date] = weekend_assigned
                # Handle abbreviated Danish month formats (e.g., "03. FEB")
                day, abbrev_month = abbrev_date_match.groups()
                month_map = {
                    'JAN': 'Januar', 'FEB': 'Februar', 'MAR': 'Marts', 'APR': 'April',
                    'MAJ': 'Maj', 'JUN': 'Juni', 'JUL': 'Juli', 'AUG': 'August',
                    'SEP': 'September', 'OKT': 'Oktober', 'NOV': 'November', 'DEC': 'December'
                }
                month_name = month_map.get(abbrev_month.upper(), abbrev_month)
                year = "2026" if month_name in ("Januar", "Februar", "Marts") else "2025"
                current_date = f"Tirsdag {day} {month_name} {year}"
                assigned = set()
                weekend_assigned = set()
                current_weekend_date = None
                if 'Ingen møde' in line:
                    current_date = None
                    continue
            
            # Check if we're entering a weekend meeting section
            if 'Weekendmødet' in line or 'Weekendopgaver' in line:
                in_weekend_section = True
            
            # Extract names for both weekly (Tuesday) and weekend (Sunday) meetings
            if current_date or current_weekend_date:
                # Extract all names from the line, handling various formats
                names_found = set()
                
                # Look for names separated by slashes (e.g., "Christopher Rüdinger/Lucas Vinzentsen")
                slash_names = re.findall(r'([A-ZÆØÅ][a-zæøåõ]+(?:\s+[A-ZÆØÅ][a-zæøåõ]+)+)(?:\s*/\s*([A-ZÆØÅ][a-zæøåõ]+(?:\s+[A-ZÆØÅ][a-zæøåõ]+)+))?', line)
                for match in slash_names:
                    if match[0]:
                        names_found.add(match[0])
                    if match[1]:
                        names_found.add(match[1])
                
                # Look for standard Danish names
                standard_names = re.findall(r'[A-ZÆØÅ][a-zæøåõ]+(?:\s+[A-ZÆØÅ][a-zæøåõ]+)+', line)
                for name in standard_names:
                    names_found.add(name)
                
                # Look for names after colons (e.g., "Bøn: Marcel Ale", "Kl. 2: Christopher Rüdinger")
                colon_names = re.findall(r':\s*([A-ZÆØÅ][a-zæøåõ]+(?:\s+[A-ZÆØÅ][a-zæøåõ]+)+)', line)
                for name in colon_names:
                    names_found.add(name)
                
                # Look for names in parentheses
                paren_names = re.findall(r'\(([A-ZÆØÅ][a-zæøåõ]+(?:\s+[A-ZÆØÅ][a-zæøåõ]+)+)\)', line)
                for name in paren_names:
                    names_found.add(name)
                
                # Match found names to members list and add to assigned set
                for pdf_name in names_found:
                    matched_member = find_matching_member(pdf_name, members_list)
                    if matched_member:
                        # If we're in a weekend section and have a weekend date, assign to weekend
                        if in_weekend_section and current_weekend_date:
                            weekend_assigned.add(matched_member)
                        # Otherwise, assign to weekly meeting (Tuesday)
                        elif current_date:
                            assigned.add(matched_member)
        
        # Save any remaining dates
        if current_date:
            meetings[current_date] = assigned
        if current_weekend_date:
            meetings[current_weekend_date] = weekend_assigned
        
        # Merge meetings from this PDF into all_meetings
        for date, assigned_people in meetings.items():
            if date in all_meetings:
                # If date already exists, merge the assigned people
                all_meetings[date].update(assigned_people)
            else:
                all_meetings[date] = assigned_people
    
    return all_meetings

def generate_schedule(members, meetings):
    dates = list(meetings.keys())
    
    # Create a mapping for Danish month names to numbers for proper sorting
    month_order = {
        'Januar': 1, 'Februar': 2, 'Marts': 3, 'April': 4, 'Maj': 5, 'Juni': 6,
        'Juli': 7, 'August': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'December': 12
    }
    
    def sort_key(date_str):
        current_year = datetime.datetime.now().year
        
        # Handle weekday dates like "Tirsdag 15 Oktober 2025", "Torsdag 09 December"
        if date_str.startswith(('Tirsdag', 'Mandag', 'Torsdag')):
            day_match = re.search(r'\d{2}', date_str)
            month_match = re.search(r'(Januar|Februar|Marts|April|Maj|Juni|Juli|August|September|Oktober|November|December)', date_str)
            year_match = re.search(r'\d{4}', date_str)
            if day_match and month_match:
                day = int(day_match.group(0))
                month = month_order.get(month_match.group(0), 0)
                year = int(year_match.group(0)) if year_match else 2025
                return (year, month, day)
        
        # Handle weekend dates like "Søndag 07 September 2025"
        elif date_str.startswith('Søndag'):
            parts = date_str.split()
            if len(parts) >= 4:
                day = int(parts[1])
                month = month_order.get(parts[2], 0)
                year = int(parts[3])
                return (year, month, day)
        
        return (0, 0, 0)
    
    dates.sort(key=sort_key)
    schedule = {}
    i = 0
    n = len(members)
    for date in dates:
        vert1 = None
        vert2 = None
        attempts = 0
        max_attempts = n * 2
        while vert1 is None and attempts < max_attempts:
            cand = members[i % n]
            i += 1
            attempts += 1
            if cand not in meetings[date]:
                vert1 = cand
        attempts = 0
        while vert2 is None and attempts < max_attempts:
            cand = members[i % n]
            i += 1
            attempts += 1
            if cand not in meetings[date]:
                vert2 = cand
        if vert1 and vert2:
            schedule[date] = (vert1, vert2)
        else:
            schedule[date] = ('No available', 'No available')
    return schedule

def create_xlsx(schedule):
    output = BytesIO()
    df = pd.DataFrame({
        'Dato': list(schedule.keys()),
        'Vært 1': [v[0] for v in schedule.values()],
        'Vært 2': [v[1] for v in schedule.values()]
    })
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Tidsplan')
    output.seek(0)
    return output

# Page configuration
st.set_page_config(
    page_title="Mødevært Schedule App",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .upload-section {
        padding: 1rem;
        margin-bottom: 1rem;
    }
    .info-box {
        background-color: #e3f2fd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2196f3;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #e8f5e8;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #4caf50;
        margin-bottom: 1rem;
    }
    .schedule-table {
        background-color: white;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .download-btn {
        background: linear-gradient(90deg, #4caf50 0%, #45a049 100%);
        color: white;
        padding: 0.75rem 1.5rem;
        border-radius: 25px;
        border: none;
        font-weight: bold;
        text-align: center;
        display: inline-block;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Main header
st.markdown("""
<div class="main-header">
    <h1>📅 Mødevært Planlægnings App</h1>
    <p>Automatisk planlægning af mødeværter med konfliktregistrering</p>
</div>
""", unsafe_allow_html=True)

# Sidebar for instructions
with st.sidebar:
    st.markdown("### 📋 Instruktioner")
    st.markdown("""
    1. **Upload Medlemsfil**: Excel-fil med godkendte mødeværter (navne fra række 3)
    2. **Upload PDF-filer**: Mødeprogrammer (flere filer understøttet)
    3. **Gennemse registrerede møder**: Kontroller hvilke opgaver der blev fundet
    4. **Generer tidsplan**: Automatisk værtstildeling med konfliktundgåelse
    5. **Download resultater**: Få din Excel-tidsplan
    """)
    
    st.markdown("### 📁 Filkrav")
    st.markdown("""
    **Excel-fil:**
    - Mødeværternes navne i første kolonne
    - Starter fra række 3
    
    **PDF-filer:**
    - **Hverdagsmøder:** Dansk format "Tirsdag 15 September"
    - **Weekendmøder:** Datoformat "DD/MM/ÅÅÅÅ" (f.eks. "07/09/2025")
    - Indeholder deltageropgaver
    """)

# Upload section
st.markdown("### Upload Filer")
with st.container():
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        uploaded_xlsx = st.file_uploader(
            '📊 **Upload Godkendte Mødeværter**', 
            type=['xlsx'],
            help="Upload Excel-fil med liste over godkendte mødeværter"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        uploaded_pdfs = st.file_uploader(
            '📄 **Upload Mødeprogrammer**', 
            type=['pdf'], 
            accept_multiple_files=True,
            help="Upload en eller flere PDF-filer med mødeprogrammer"
        )
        st.markdown('</div>', unsafe_allow_html=True)

if uploaded_xlsx and uploaded_pdfs:
    members = parse_members(uploaded_xlsx)
    
    # Show uploaded files info
    st.markdown('<div class="info-box">', unsafe_allow_html=True)
    st.markdown(f"**📄 Uploadede Filer:**")
    st.markdown(f"- **Medlemsfil:** {uploaded_xlsx.name}")
    st.markdown(f"- **PDF-filer:** {len(uploaded_pdfs)} fil(er)")
    for i, pdf in enumerate(uploaded_pdfs, 1):
        st.markdown(f"  {i}. {pdf.name}")
    st.markdown('</div>', unsafe_allow_html=True)
    
    meetings = parse_program(uploaded_pdfs, members)
    if meetings:
        # Success message
        st.markdown('<div class="success-box">', unsafe_allow_html=True)
        st.markdown(f"**✅ Succes!** {len(meetings)} møde(r) registreret og behandlet")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Meeting details in expandable section
        with st.expander("📋 **Se Registrerede Møder og Opgaver**", expanded=False):
            # Use the same sorting logic as in generate_schedule
            def display_sort_key(date_str):
                month_order = {
                    'Januar': 1, 'Februar': 2, 'Marts': 3, 'April': 4, 'Maj': 5, 'Juni': 6,
                    'Juli': 7, 'August': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'December': 12
                }
                
                # Handle weekday dates like "Tirsdag 15 Oktober 2025", "Torsdag 09 December"
                if date_str.startswith(('Tirsdag', 'Mandag', 'Torsdag')):
                    day_match = re.search(r'\d{2}', date_str)
                    month_match = re.search(r'(Januar|Februar|Marts|April|Maj|Juni|Juli|August|September|Oktober|November|December)', date_str)
                    year_match = re.search(r'\d{4}', date_str)
                    if day_match and month_match:
                        day = int(day_match.group(0))
                        month = month_order.get(month_match.group(0), 0)
                        year = int(year_match.group(0)) if year_match else 2025
                        return (year, month, day)
                
                # Handle weekend dates like "Søndag 07 September 2025"
                elif date_str.startswith('Søndag'):
                    parts = date_str.split()
                    if len(parts) >= 4:
                        day = int(parts[1])
                        month = month_order.get(parts[2], 0)
                        year = int(parts[3])
                        return (year, month, day)
                
                return (0, 0, 0)
            
            for date in sorted(meetings.keys(), key=display_sort_key):
                assigned_people = meetings[date]
                if assigned_people:
                    st.markdown(f"**{date}:** {', '.join(sorted(assigned_people))}")
                else:
                    st.markdown(f"**{date}:** *(ingen opgaver registreret)*")
        
        # Generate and display schedule
        schedule = generate_schedule(members, meetings)
        
        st.markdown("### 📅 Genereret Tidsplan")
        st.markdown('<div class="schedule-table">', unsafe_allow_html=True)
        
        # Create a styled dataframe
        df_schedule = pd.DataFrame({
            'Dato': list(schedule.keys()),
            'Vært 1': [v[0] for v in schedule.values()],
            'Vært 2': [v[1] for v in schedule.values()]
        })
        
        # Display with better formatting
        st.dataframe(
            df_schedule,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Dato": st.column_config.TextColumn("📅 Dato", width="medium"),
                "Vært 1": st.column_config.TextColumn("👤 Vært 1", width="medium"),
                "Vært 2": st.column_config.TextColumn("👤 Vært 2", width="medium")
            }
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Download section
        st.markdown("### 💾 Download Resultater")
        output = create_xlsx(schedule)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.download_button(
                label='📥 Download Tidsplan (XLSX)',
                data=output,
                file_name='modevart_tidsplan.xlsx',
                mime='application/vnd.ms-excel',
                use_container_width=True
            )
        
        # Summary statistics
        st.markdown("### 📊 Oversigt")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Møder", len(schedule))
        with col2:
            st.metric("Total Værter Tildelt", len(schedule) * 2)
        with col3:
            st.metric("Tilgængelige Mødeværter", len(members))
        with col4:
            conflicts_avoided = sum(1 for date, assigned in meetings.items() if len(assigned) > 0)
            st.metric("Konflikter Undgået", conflicts_avoided)
            
    else:
        st.error('❌ Ingen møder fundet i de uploadede PDF-filer. Kontroller venligst dine filer og prøv igen.')
