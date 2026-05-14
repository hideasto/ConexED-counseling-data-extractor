# ConexED Counseling Appointment Data Extractor

A Python script that reads counseling appointment records exported from **ConexED** (Master Appointment Report) and outputs a structured, analysis-ready Excel file.

---

## Background

ConexED is an appointment scheduling platform used by counseling departments at community colleges. When counselors export appointment data, the report comes as a PDF called the **Master Appointment Report**. While the data is all there, it is not in a format that can be directly loaded into Power BI, Excel pivot tables, or any analytical tool without significant manual cleanup.

This script handles the extraction automatically — reading the PDF page by page, parsing each appointment row, standardizing the reason codes, and saving everything to a clean Excel file ready for analysis.

This project is demonstrated using fully anonymized mock data (cartoon character students, actor counselors, musician schedulers). No real student data is included in this repository.

---

## ConexED vs SARS — Key Difference

This script is specifically for **ConexED** data. If your campus uses **SARS** to track counseling appointments, see the separate SARS extraction repository. The two systems produce very different report formats:

| Feature | SARS | ConexED |
|---|---|---|
| Report name | Short Name History Report | Master Appointment Report |
| Counselor location | Page header | Moderator column |
| Reason format | Short codes (`CC`, `VA-PHONE`) | Full phrases (`Student Educational Plan`) |
| Drop-in detection | Source column text | Dedicated Drop-In column (Y/N) |
| Orientation | Portrait | Landscape |

---

## What the Script Does

1. **Reads the PDF** page by page using `pdfplumber`
2. **Extracts all 16 columns** from each appointment row
3. **Cleans line-break artifacts** — pdfplumber sometimes splits text mid-word across lines due to narrow columns; the script rejoins these correctly
4. **Parses the Scheduled Start** date and time into separate clean fields
5. **Converts Meeting Length** from `HH:MM:SS` format to total minutes
6. **Determines appointment type** (Drop-in vs The Grid) from the Drop-In column
7. **Detects modality** (In-Person, Zoom, Phone) from the Type column
8. **Matches reason codes** against a standardized dictionary and stores each match in its own column (`Reason_1` through `Reason_20`)
9. **Adds calculated columns**: Term code, Term label, Intersession flag
10. **Saves everything** to an Excel file in the same folder as the PDF

---

## Output Columns

| Column | Description |
|---|---|
| Source | The Grid or Drop-in |
| Appointment Type (Source) | Appointment or Drop-In |
| Attendance Status | Attended or Not Attended |
| INTERSESSION | 1 if date falls in January intersession, 0 otherwise |
| Counselor | Moderator name from the ConexED report |
| Student ID | Attendee SIS ID |
| Student Name | Attendee name |
| Date | Appointment date (YYYY-MM-DD) |
| Term | Term code (e.g. 22253) |
| Term Recode | Human-readable term (e.g. Spring 2025) |
| Time | Appointment time |
| Duration | Meeting length in minutes |
| Total_Reasons | Number of reason codes matched |
| Modality | In-Person, Zoom, or Phone |
| Reason Code(s) / Comments | Full reported reason text |
| Reason_1 … Reason_20 | Individual standardized reason codes |
| ScheduledStart | Original raw Scheduled Start value |

---

## How Reason Code Matching Works

ConexED stores reasons as full phrases in the **Reported Reason(s)** column, for example:

```
Nextup Counseling Appointment, Other reason personal, Transfer to 4-year University
```

The script matches each phrase against a dictionary of standardized labels (`STANDARDIZED_REASONS`) at the top of the file. Each match is stored in its own column:

```
Reason_1 = 'Nextup Counseling Appointment (EOPS)'
Reason_2 = 'Other reason personal (EOPS)'
Reason_3 = 'Transfer Advising (EOPS)'
Total_Reasons = 3
```

### Why Reason_1 through Reason_20?

A single appointment can have multiple reasons. Splitting each into its own column allows you to **unpivot** them in Power BI — giving you one row per reason per appointment — so you can accurately count how often each reason code appears and build breakdowns by counselor, term, or modality.

### Updating reason codes for your campus

The `STANDARDIZED_REASONS` dictionary maps standardized labels to the raw text variants that may appear in your PDF. Each key is the label that appears in the output Excel; each value is a list of raw text variants that should map to that label.

```python
STANDARDIZED_REASONS = {
    'Student Educational Plan (EOPS)': ['Student Educational Plan'],
    'Transfer Advising (EOPS)':        ['Transfer Advising'],
    # Add your campus-specific reasons here
}
```

Before running this script on your real PDF, open a few pages and review the Reported Reason(s) column. If you see reason phrases that are not in the dictionary, add them. The script will silently ignore any reason text that does not match an entry in the dictionary.

---

## How Line Breaks Are Handled

ConexED PDFs are printed in landscape orientation, but even so, some cells contain text that wraps across multiple lines within a cell. `pdfplumber` returns these as strings with `\n` characters inside them.

The script handles three types of line breaks differently:

- **Names** (`clean_name`) — detects whether the break is mid-word (no space before the newline) and joins directly, or at a word boundary and adds a space. Example: `'Denzel Wash\nington'` → `'Denzel Washington'`
- **Reason text** (`clean_text_with_linebreaks`) — joins lines with spaces and removes trailing commas, preserving the full phrase for matching
- **Dates** (`split_scheduled_start`) — joins lines with no separator first, then uses regex to extract the date and time components regardless of how they were split

---

## Known Limitation — Records with No Reason Match

Some records will show `Total_Reasons = 0`. This happens when:
- The Reported Reason(s) cell is blank
- The reason text uses phrasing not in the `STANDARDIZED_REASONS` dictionary

If you see a high number of zero-reason records, open your PDF and check what phrases appear in the Reported Reason(s) column, then add any missing ones to the dictionary.

---

## Requirements

```
pandas
pdfplumber
openpyxl
```

Install with:

```bash
pip install pandas pdfplumber openpyxl
```

---

## How to Run

1. Clone or download this repository
2. Place your ConexED PDF in the same folder as the script
3. Open `ConexED_extraction.py` and update these two lines in `main()`:

```python
pdf_directory = r"C:\path\to\your\folder"       # folder containing the PDF
pdf_filename  = "your_conexed_report.pdf"        # exact filename
```

4. Review and update `STANDARDIZED_REASONS` at the top of the script to match your campus reason codes
5. Double-click the script, or run from the command line:

```bash
python ConexED_extraction.py
```

6. The output Excel file (`ConexED_Cleaned_Data.xlsx`) will appear in the same folder as the PDF

---

## Files in This Repository

| File | Description |
|---|---|
| `ConexED_extraction.py` | Main extraction script |
| `Mock_ConexED_Counseling_Data.pdf` | Sample landscape PDF with anonymized mock data for testing |
| `README.md` | This file |

---

## Mock Data Note

The sample PDF included in this repository uses:
- **Cartoon characters** as student names and emails (e.g. `donald.duck@mouseclub.edu`)
- **Alphanumeric IDs** like `N7N8XJWH` instead of real student IDs
- **Famous actors** as counselors (Denzel Washington, Meryl Streep, Tom Hanks, etc.)
- **Famous musicians** in the Scheduled By column (Elton John, Madonna, Aretha Franklin, etc.)
- **Impossible dates** in Scheduled By (e.g. `Jul 71, 3041`) to make it obvious the data is fake
- **Real ConexED reason phrases** so the matching logic can be tested end to end

No real student records, counselor names, or institutional data are included.

---

## Adapting for Other Counseling Programs

This script was built for EOPS/CARE/NEXT UP ConexED data. If your campus uses ConexED for other programs (CalWORKs, Disability Services, etc.), the extraction logic is the same — only `STANDARDIZED_REASONS` needs to be updated to reflect that program's reason phrases.
