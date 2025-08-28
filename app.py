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

def parse_program(uploaded_files, members_list):
    all_meetings = {}
    
    for uploaded_file in uploaded_files:
        with pdfplumber.open(uploaded_file) as pdf:
            text = '\n'.join(page.extract_text() for page in pdf.pages if page.extract_text())
        meetings = {}
        current_date = None
        assigned = set()
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            # Look for weekday meeting dates (e.g., "Tirsdag 15 September")
            weekday_date_match = re.match(r'(Tirsdag|Mandag) \d{2} (September|Oktober|November|December|Januar|Februar|Marts|April|Maj|Juni|Juli|August)', line)
            
            # Look for weekend meeting dates (e.g., "07/09/2025")
            weekend_date_match = re.match(r'(\d{2})/(\d{2})/(\d{4})', line)
            
            if weekday_date_match:
                if current_date:
                    meetings[current_date] = assigned
                # Add year to weekday dates for consistency
                weekday_date = weekday_date_match.group(0)
                current_date = f"{weekday_date} 2025"
                assigned = set()
                if 'Ingen møde' in line:
                    current_date = None
                    continue
            elif weekend_date_match:
                if current_date:
                    meetings[current_date] = assigned
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
            if current_date:
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
                        assigned.add(matched_member)
        
        if current_date:
            meetings[current_date] = assigned
        
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
        
        # Handle weekday dates like "Tirsdag 15 Oktober 2025"
        if date_str.startswith(('Tirsdag', 'Mandag')):
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
                
                # Handle weekday dates like "Tirsdag 15 Oktober 2025"
                if date_str.startswith(('Tirsdag', 'Mandag')):
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
