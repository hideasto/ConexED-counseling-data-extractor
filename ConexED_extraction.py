"""
ConexED Counseling Appointment Data Extractor
=============================================
Reads a ConexED Master Appointment Report PDF, extracts all appointment
records, and outputs a structured, analysis-ready Excel file.

Dependencies:
    pip install pandas pdfplumber openpyxl

Usage:
    1. Set pdf_directory and pdf_filename in main() to point to your PDF.
    2. Double-click the script or run: python ConexED_extraction.py
    3. The output Excel file is saved in the same folder as the PDF.
"""

import os
import pandas as pd
import pdfplumber
import warnings
import re
from datetime import datetime

warnings.filterwarnings("ignore")


# ── Reason code mapping ────────────────────────────────────────────────────
# Maps standardized output labels to the raw text variants that may appear
# in the Reported Reason(s) column of the ConexED PDF.
# Update this dictionary to match the reason codes used at your campus.

STANDARDIZED_REASONS = {
    'Academic Disq (EOPS)':
        ['Academic Disq', 'Academic Disqualification'],
    'Career Counseling (EOPS)':
        ['Career Counseling'],
    'Counselor Contact (EOPS)':
        ['Counselor Contact'],
    'Drop In (EOPS)':
        ['Drop In'],
    'Enrolled in Nextup Program (EOPS)':
        ['Enrolled in Nextup Program', 'Enrolled in NextUp Program'],
    'EOPS Counseling Appointment (EOPS)':
        ['EOPS Counseling Appointment', 'EOPs Counseling Appointment'],
    'EOPS Non-Compliance (EOPS)':
        ['EOPS Non-Compliance', 'EOPs Non-Compliance'],
    'EOPS Orientation Session (EOPS ORIEN) (VAR)':
        ['EOPS Orientation Session'],
    'EOPS Technician Appointment (non-academic) (EOPS)':
        ['EOPS Technician Appointment', 'EOPs Technician Appointment'],
    'EOPS Workshop Session (EOPS WKS) (VAR)':
        ['EOPS Workshop Session'],
    'Financial Aid Appeal (EOPS)':
        ['Financial Aid Appeal'],
    'Group Counseling Online Workshop (EOPS)':
        ['Group Counseling Online Workshop'],
    'High Unit Major Petition (EOPS)':
        ['High Unit Major Petition'],
    'Nextup Counseling Appointment (EOPS)':
        ['Nextup Counseling Appointment', 'NextUp Counseling Appointment'],
    'Other reason personal (EOPS)':
        ['Other reason personal'],
    'Progress Report (EOPS)':
        ['Progress Report'],
    'Student Educational Plan (EOPS)':
        ['Student Educational Plan'],
    'Transfer Advising (EOPS)':
        ['Transfer Advising'],
}


# ── Text cleaning ──────────────────────────────────────────────────────────

def clean_name(text):
    """
    Clean a person's name that has been split mid-word by PDF line wrapping.
    Example: 'Denzel Wash\\nington' → 'Denzel Washington'
             'Anthony\\nHopkins'    → 'Anthony Hopkins'
             'Duck, Louie\\nC'      → 'Duck, Louie C'

    Uses the raw text before stripping to detect whether the break is
    mid-word (no space before \\n) or a word boundary (space before \\n).
    """
    if not text:
        return ""

    # Split on newlines but keep track of whether each break was mid-word
    raw_lines = str(text).split('\n')
    result = ""
    for i, raw_line in enumerate(raw_lines):
        stripped = raw_line.strip()
        if not stripped:
            continue
        if i == 0:
            result = stripped
        else:
            # Check the END of the previous raw line (before stripping)
            prev_raw = raw_lines[i - 1]
            prev_ends_mid_word = (prev_raw and not prev_raw[-1].isspace()
                                  and prev_raw[-1].isalpha())
            # Check the START of this line
            curr_starts_mid_word = stripped and stripped[0].isalpha()

            if prev_ends_mid_word and curr_starts_mid_word:
                result += stripped   # mid-word: join directly, no space
            else:
                result += " " + stripped  # word boundary: add space
    return result.strip()


def clean_text_with_linebreaks(text):
    """Normalize text that contains awkward line breaks from PDF extraction."""
    if not text:
        return ""
    lines = [l.strip().rstrip(',') for l in str(text).split('\n') if l.strip()]
    result = " ".join(lines)
    result = re.sub(r',\s*,', ', ', result)
    result = re.sub(r'\s+', ' ', result)
    result = re.sub(r',\s*$', '', result)
    return result


def clean_duration(duration_text):
    """
    Convert duration text to total minutes as a string.
    Handles HH:MM:SS format (e.g. '00:38:52' → '38') and plain minutes.
    """
    if not duration_text:
        return ""

    text = re.sub(r'\s+', ' ', str(duration_text).replace('\n', ' ')).strip()
    text = re.sub(r'[^\d:]', '', text)

    if text and ':' not in text:
        try:
            minutes = int(text)
            return str(minutes // 60 if minutes > 480 else minutes)
        except:
            pass

    if ':' in text:
        parts = text.split(':')
        try:
            total = 0
            if parts[0]:
                total += int(parts[0]) * 60
            if len(parts) > 1 and parts[1]:
                total += int(parts[1])
            if len(parts) > 2 and parts[2]:
                if int(parts[2]) >= 30:
                    total += 1
            return str(total)
        except:
            pass

    nums = re.findall(r'\d+', duration_text)
    if nums:
        n = int(nums[0])
        return str(n // 60 if n > 480 else n)

    return ""


# ── Reason parsing ─────────────────────────────────────────────────────────

def parse_reasons_with_mapping(reason_text):
    """
    Scan the Reported Reason(s) text against STANDARDIZED_REASONS and return:
      - count of matched reasons
      - list of matched standardized labels
      - cleaned original text
      - whether the text indicates a drop-in / walk-in appointment
    """
    if not reason_text:
        return 0, [], "", False

    cleaned   = clean_text_with_linebreaks(reason_text)
    lower     = cleaned.lower()
    found     = []
    has_dropin = any(t in lower for t in
                     ['walk in','walk-in','drop in','drop-in','dropin','walkin'])

    for standard, variants in STANDARDIZED_REASONS.items():
        for variant in variants:
            words = variant.lower().split()
            if all(w in lower for w in words):
                found.append(standard)
                break

    return len(found), found, cleaned, has_dropin


# ── Appointment type determination ─────────────────────────────────────────

def determine_appointment_type(row, reported_reasons, scheduled_reasons,
                                has_dropin_reason, duration, drop_in_raw):
    """
    Classify appointment as Drop-in or The Grid based solely on the
    Drop-In column (Y = Drop-in, N = The Grid). This column is the most
    reliable signal in ConexED and avoids false positives from reason text.
    """
    if str(drop_in_raw).strip().upper() == "Y":
        return "Drop-in", "Drop-In"

    return "The Grid", "Appointment"


# ── Date / time splitting ──────────────────────────────────────────────────

def split_scheduled_start(value):
    """
    Parse the Scheduled Start cell into separate date and time strings.
    Handles multi-line values like '2025-07\\n-08 10:1\\n5 AM' by first
    joining all lines into a single clean string before parsing.
    """
    if not value:
        return "", ""

    # Join all lines into one string, removing line breaks
    lines = [l.strip() for l in str(value).split('\n') if l.strip()]
    text  = "".join(lines)          # join with NO space — handles split numbers
    text  = re.sub(r'\s+', ' ', text).strip()

    # Try direct parse
    try:
        dt = datetime.strptime(text, "%Y-%m-%d %I:%M %p")
        return dt.strftime("%Y-%m-%d"), dt.strftime("%I:%M %p")
    except:
        pass

    # Fallback: extract date and time components with regex
    date_m = re.search(r'(\d{4}-\d{1,2}-\d{1,2})', text)
    time_m = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', text, re.IGNORECASE)

    if date_m and time_m:
        try:
            date_part = datetime.strptime(
                date_m.group(1), "%Y-%m-%d").strftime("%Y-%m-%d")
            time_part = datetime.strptime(
                f"{time_m.group(1)} {time_m.group(2).upper()}", "%I:%M %p"
            ).strftime("%I:%M %p")
            return date_part, time_part
        except:
            pass

    return "", ""


# ── Term and intersession helpers ──────────────────────────────────────────

def get_intersession(date_str):
    """Return intersession label for January dates, otherwise empty string."""
    if not date_str:
        return ""
    try:
        dt   = datetime.strptime(date_str, "%Y-%m-%d")
        year = dt.year
        month= dt.month
        day  = dt.day

        windows = {
            2020:(2,29), 2021:(4,30), 2022:(3,29),
            2023:(3,28), 2024:(2,27), 2025:(6,31),
        }

        if month == 1 and year in windows:
            start, end = windows[year]
            if start <= day <= end:
                return f"Intersession {year}"
        return ""
    except:
        return ""


def get_term_codes(date_str):
    """Convert YYYY-MM-DD to term code and human-readable term label."""
    if not date_str:
        return "", ""
    try:
        dt    = datetime.strptime(date_str, "%Y-%m-%d")
        year  = dt.year
        month = dt.month
        day   = dt.day

        if 1 <= month <= 5:
            digit, name = 3, "Spring"
        elif month in (6, 7):
            digit, name = 5, "Summer"
        elif month == 8:
            digit, name = (5, "Summer") if day <= 14 else (7, "Fall")
        elif 9 <= month <= 12:
            digit, name = 7, "Fall"
        else:
            digit, name = 0, "Unknown"

        return f"2{str(year)[-2:]}{digit}", f"{name} {year}"
    except:
        return "", ""


# ── PDF extraction ─────────────────────────────────────────────────────────

def extract_pdf_data(pdf_path):
    """
    Extract all table rows from the ConexED PDF.
    Skips header rows (identified by 'moderator' in the first cell).
    Returns a list of raw row lists.
    """
    print(f"Reading PDF: {os.path.basename(pdf_path)}")
    all_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"Total pages: {total_pages}")

        for page_num, page in enumerate(pdf.pages, 1):
            print(f"  Processing page {page_num}/{total_pages}")
            tables = page.extract_tables({
                "vertical_strategy":   "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance":      3,
            })

            for table in tables:
                for row in table:
                    cleaned = [str(c).strip() if c else "" for c in row]
                    if not any(cleaned):
                        continue
                    if cleaned[0].lower() == "moderator":
                        continue
                    all_rows.append(cleaned)

    print(f"Extracted {len(all_rows)} rows")
    return all_rows


# ── Row processing ─────────────────────────────────────────────────────────

def process_pdf_to_final(pdf_path):
    """
    Extract raw rows from the PDF and transform each into a structured
    output record matching the final Excel column layout.

    ConexED column order (0-indexed):
      0  Moderator          8  Staff Review Completed
      1  Attendee           9  Type
      2  Attendee SIS ID   10  Drop-In
      3  Attendee Email    11  Group
      4  Phone Number      12  Location
      5  Scheduled By      13  Meeting Length
      6  Scheduled Reason  14  Scheduled Start
      7  Reported Reason   15  Date Scheduled
    """
    raw_data = extract_pdf_data(pdf_path)

    if not raw_data:
        print("ERROR: No data extracted from PDF.")
        return []

    print(f"\nFirst row (column check):")
    for i, col in enumerate(raw_data[0]):
        print(f"  Column {i}: {repr(col)}")
    print()

    processed = []

    for row_idx, row in enumerate(raw_data):
        if len(row) < 15:
            continue

        # Skip any stray header rows
        if row_idx == 0 and any(
            col.lower() in ['counselor','student name','student id']
            for col in row[:3]
        ):
            continue

        counselor          = clean_name(row[0])
        student_name       = clean_name(row[1])
        student_id         = clean_name(row[2])
        scheduled_reasons  = clean_text_with_linebreaks(row[6] if len(row) > 6  else "")
        reported_reasons   = clean_text_with_linebreaks(row[7] if len(row) > 7  else "")
        drop_in_raw        = clean_text_with_linebreaks(row[10] if len(row) > 10 else "")
        meeting_length_raw = clean_text_with_linebreaks(row[13] if len(row) > 13 else "")
        scheduled_start    = clean_text_with_linebreaks(row[14] if len(row) > 14 else "")

        date_value, time_value = split_scheduled_start(scheduled_start)
        duration = clean_duration(meeting_length_raw)

        reason_count, std_reasons, original_reasons, has_dropin = \
            parse_reasons_with_mapping(reported_reasons)

        # Attendance: ConexED marks no-shows in the row text
        attendance = ("Not Attended"
                      if "NO SHOW" in " ".join(row).upper()
                      else "Attended")

        # Modality from Type column
        appt_type_raw = row[9] if len(row) > 9 else ""
        lower_type    = str(appt_type_raw).lower()
        if "online" in lower_type or "video" in lower_type:
            modality = "Zoom"
        elif "phone" in lower_type:
            modality = "Phone"
        else:
            modality = "In-Person"

        intersession_value = "1" if get_intersession(date_value) else "0"
        term_code, term_recode = get_term_codes(date_value)

        source, appt_type = determine_appointment_type(
            row, reported_reasons, scheduled_reasons,
            has_dropin, duration, drop_in_raw
        )

        # Debug first 3 rows
        if row_idx < 3:
            print(f"Row {row_idx}:")
            print(f"  Counselor: {counselor}")
            print(f"  Student:   {student_name}")
            print(f"  Duration raw: {repr(meeting_length_raw)} → {duration} min")
            print(f"  Drop-In col: {repr(drop_in_raw)} → Source: {source}")
            print(f"  Date/Time: {date_value} {time_value}")
            print(f"  Modality:  {modality}")
            print()

        out_row = {
            "Source":                    source,
            "Appointment Type (Source)": appt_type,
            "Attendance Status":         attendance,
            "INTERSESSION":              intersession_value,
            "Counselor":                 counselor,
            "Student ID":                student_id,
            "Student Name":              student_name,
            "ScheduledStart":            scheduled_start,
            "Date":                      date_value,
            "Term":                      term_code,
            "Term Recode":               term_recode,
            "Time":                      time_value,
            "Duration":                  duration,
            "Modality":                  modality,
            "Total_Reasons":             reason_count,
            "Reason Code(s) / Comments": original_reasons,
        }

        for j in range(1, 21):
            out_row[f"Reason_{j}"] = (std_reasons[j-1]
                                       if j <= len(std_reasons) else "")

        processed.append(out_row)

    print(f"\nProcessed {len(processed)} appointment rows")
    return processed


# ── Entry point ────────────────────────────────────────────────────────────

def main():
    print("="*80)
    print("CONEXED COUNSELING APPOINTMENT DATA EXTRACTOR")
    print("="*80)

    # ── Update these two lines to point to your PDF ──
    pdf_directory = r"C:\Users\hasto\Desktop\Counseling ConexEd"
    pdf_filename  = "Mock_ConexED_Counseling_Data.pdf"
    pdf_path      = os.path.join(pdf_directory, pdf_filename)

    if not os.path.exists(pdf_path):
        print(f"\nERROR: File not found:\n  {pdf_path}")
        print("\nPDF files in that directory:")
        try:
            for f in os.listdir(pdf_directory):
                if f.lower().endswith('.pdf'):
                    print(f"  - {f}")
        except Exception as e:
            print(f"  Could not list directory: {e}")
        input("\nPress Enter to exit...")
        return

    print(f"\nProcessing: {pdf_path}")
    rows = process_pdf_to_final(pdf_path)

    if not rows:
        print("ERROR: No data to process.")
        input("\nPress Enter to exit...")
        return

    df = pd.DataFrame(rows)

    column_order = [
        'Source', 'Appointment Type (Source)', 'Attendance Status', 'INTERSESSION',
        'Counselor', 'Student ID', 'Student Name', 'Date', 'Term', 'Term Recode',
        'Time', 'Duration', 'Total_Reasons', 'Modality', 'Reason Code(s) / Comments',
        *[f'Reason_{i}' for i in range(1, 21)],
        'ScheduledStart',
    ]
    column_order = [c for c in column_order if c in df.columns]
    df = df[column_order]

    output_file = os.path.join(pdf_directory, "ConexED_Cleaned_Data.xlsx")

    try:
        df.to_excel(output_file, index=False)
        print(f"\nSUCCESS — file saved to:\n  {output_file}")
        print(f"Total records: {len(df):,}")

        print("\n" + "="*80)
        print("DATA SUMMARY")
        print("="*80)

        for label, series in [
            ("Source",            df['Source']),
            ("Counselor",         df['Counselor']),
            ("Attendance Status", df['Attendance Status']),
            ("INTERSESSION",      df['INTERSESSION']),
            ("Modality",          df['Modality']),
        ]:
            print(f"\n{label}:")
            for val, cnt in series.value_counts().items():
                print(f"  '{val}': {cnt:,}")

        if 'Total_Reasons' in df.columns:
            print("\nRecords by reason count:")
            print(df['Total_Reasons'].value_counts().sort_index().to_string())

            reason_counts = {}
            for i in range(1, 21):
                col = f'Reason_{i}'
                if col in df.columns:
                    for val, cnt in df[col].value_counts().items():
                        if val:
                            reason_counts[val] = reason_counts.get(val, 0) + cnt
            print("\nTop 10 reason codes:")
            for code, cnt in sorted(reason_counts.items(),
                                    key=lambda x: x[1], reverse=True)[:10]:
                print(f"  {code}: {cnt:,}")

        try:
            df['Duration_num'] = pd.to_numeric(df['Duration'], errors='coerce')
            print(f"\nDuration (minutes):")
            print(f"  Average: {df['Duration_num'].mean():.1f}")
            print(f"  Min:     {df['Duration_num'].min():.0f}")
            print(f"  Max:     {df['Duration_num'].max():.0f}")
        except:
            pass

        print(f"\nUnique students:   {df['Student ID'].nunique():,}")
        print(f"Unique counselors: {df['Counselor'].nunique():,}")
        print(f"Date range:        {df['Date'].min()} to {df['Date'].max()}")
        print("\n" + "="*80)

    except Exception as e:
        print(f"\nError saving Excel file: {e}")

    input("\nProcessing complete. Press Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + "="*80)
        print("UNEXPECTED ERROR:")
        print(f"  {type(e).__name__}: {e}")
        print("\nCommon fixes:")
        print("  - Install missing libraries:")
        print("      py -m pip install pandas pdfplumber openpyxl")
        print("  - Check that pdf_directory and pdf_filename are correct in main()")
        print("="*80)
        input("\nPress Enter to close...")
