"""Open the report in Microsoft Word, refresh the table of contents and all fields, then export a PDF.

Usage (from the repository root, Windows with Word installed):
    python scripts/export_pdf.py
"""
from pathlib import Path

import win32com.client

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "report" / "Student_Depression_Report.docx"
PDF = DOCX.with_suffix(".pdf")
WD_FORMAT_PDF = 17

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(str(DOCX))
    # Update twice: the first pass fills the TOC, the second fixes page numbers shifted by the TOC itself.
    for _ in range(2):
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
    doc.Save()
    doc.ExportAsFixedFormat(str(PDF), WD_FORMAT_PDF)
    pages = doc.ComputeStatistics(2)  # wdStatisticPages
    doc.Close(False)
    print(f"Saved {PDF.relative_to(ROOT)} ({pages} pages)")
finally:
    word.Quit()
