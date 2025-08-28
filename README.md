# Mødevært Schedule App

A Streamlit application for generating meeting host schedules. The app processes Excel files containing approved members and PDF files with meeting schedules to automatically assign hosts for each meeting.

## Features

- Upload Excel files with approved members
- Upload PDF files with meeting schedules
- Automatic host assignment based on availability
- Download generated schedules as Excel files
- Danish language support for meeting dates and names

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

1. **Start the Streamlit app:**
   ```bash
   streamlit run app.py
   ```

2. **Open your browser** and navigate to the URL shown in the terminal (usually `http://localhost:8501`)

## How to Use

1. **Upload approved members file:** Use the first file uploader to upload an Excel (.xlsx) file containing the list of approved members. The app expects the member names to be in the first column starting from row 3.

2. **Upload meeting schedule PDF:** Use the second file uploader to upload a PDF file containing the meeting schedule. The app looks for Danish date patterns like "Tirsdag 15 September" or "Mandag 20 September".

3. **Generate schedule:** Once both files are uploaded, the app will automatically process them and display the generated schedule with assigned hosts.

4. **Download results:** Click the "Download Schedule XLSX" button to download the generated schedule as an Excel file.

## File Format Requirements

### Excel File (Members)
- First column should contain member names
- Names should start from row 3 (first two rows are ignored)
- Names should be in Danish format (e.g., "Jens Hansen", "Marie Jensen")

### PDF File (Meeting Schedule)
- Should contain meeting dates in Danish format: "Tirsdag 15 September" or "Mandag 20 September"
- Should contain participant names in Danish format
- Lines with "Ingen møde" (no meeting) will be skipped

## Dependencies

- `streamlit`: Web application framework
- `pandas`: Data manipulation and Excel file handling
- `pdfplumber`: PDF text extraction
- `xlsxwriter`: Excel file creation
- `openpyxl`: Excel file reading support
