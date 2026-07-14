# Research Paper Metadata Extractor

Extracts **Paper Title, Author Names and Author Emails** from research paper PDFs into one Excel file — one row per author email, with paper info repeated on every row.

Output columns: `Paper ID / Paper Name | Paper Title | All Author Names | Author Email`

## Requirements
- Python with `PyMuPDF` and `openpyxl` (already installed on this machine)

## Web Dashboard

```powershell
python app.py
```
Then open **http://127.0.0.1:5075** in your browser:
- **Drag & drop PDFs** → metadata is extracted and shown in a live table (one row per email), with status pills flagging papers missing a title, authors, or emails.
- **Download Excel** → exports everything in the required 4-column format.
- **Check an existing .xlsx** → drop a metadata sheet to see blank/invalid emails, duplicate emails within a paper, and inconsistent title/author rows.
- Stats bar shows papers processed, emails found, Excel rows, and files needing review. Remove individual papers with ✕ or clear everything.

## Command-line Usage

Put all your paper PDFs (e.g. `868.pdf`, `873.pdf`, ...) into one folder, then:

```powershell
# Extract everything into a new Excel file
python extract_metadata.py "C:\path\to\pdf_folder"

# Custom output name
python extract_metadata.py "C:\path\to\pdf_folder" -o MyPapers.xlsx

# Add new papers to an existing sheet (already-listed PDFs are skipped)
python extract_metadata.py "C:\path\to\pdf_folder" --append Research_Paper_Metadata.xlsx

# Validate an existing sheet (blank/invalid emails, duplicate emails per paper,
# inconsistent titles/authors, header mismatch)
python extract_metadata.py --check "C:\Users\DELL\Downloads\PDF NOrmal.xlsx"
```

## How it works
- **Title**: largest-font text block in the top half of page 1 (multi-line titles are joined).
- **Authors**: name-like lines between the title and the Abstract/affiliations.
- **Emails**: regex scan of the first 2 pages; grouped forms like `{a, b}@uni.edu` are expanded into separate addresses; duplicates within a paper are removed; order preserved.
- **Missing email**: the row is still written with a blank Author Email cell, and the file is flagged in the console for manual review.
- **Scanned PDFs** (no text layer) are flagged with an error — they need OCR first.

## Notes
- Filenames are used exactly as-is for Paper ID.
- Emails and names are never modified, translated, or guessed.
- Console output flags every file where title/authors/emails could not be found, so you know exactly what to review manually — nothing is silently dropped.
