#!/usr/bin/env python3
"""
gst_pdf_to_excel.py - Convert GSTR-9, GSTR-9C, and BO Profile (GST risk
intelligence) PDF reports into clean, properly-sectioned Excel workbooks.

Works identically on a PC and inside Pydroid 3 on Android.

DESKTOP / TERMINAL USAGE (unchanged from the PC version)
    python3 gst_pdf_to_excel.py <input.pdf> [output.xlsx]
    python3 gst_pdf_to_excel.py <folder_of_pdfs> [output_folder]

PYDROID 3 / "TAP THE RUN BUTTON" USAGE
    Pydroid's Run button launches the script with no command-line
    arguments, so this script also supports a zero-argument mode:
    just tap Run. It will automatically create two folders next to
    this script:

        PDF_Input/     <- put your GSTR-9 / GSTR-9C / BO Profile PDFs here
        Excel_Output/  <- converted .xlsx workbooks appear here

    Use Pydroid's built-in file browser to drop PDFs into PDF_Input,
    then tap Run again. Every PDF in that folder is converted in one
    pass; a summary is printed for each file.

The document type (GSTR-9C, GSTR-9, or BO Profile) is auto-detected from the
PDF's own text, so this one script works for any future GSTIN/filing you feed
it - no need to touch the code per client.

DESIGN PRINCIPLE
    A worksheet holds the COMPLETE data for exactly one logical heading /
    section of the source PDF. A new heading always starts a new worksheet.
    No section is ever split across two worksheets, and the split points are
    the PDF's own structural markers (item numbers, "Pt." labels, or known
    field/column-header names) - never a fixed row count or page boundary.

Requires: pdfplumber, openpyxl
    On Pydroid 3: menu -> Pip -> type "openpyxl", install, then type
    "pdfplumber", install. See the setup notes shipped alongside this
    file for troubleshooting slow/compiled-dependency installs on Android.
"""
import os
import sys
import traceback
import pdfplumber

import parse_gstr9
import parse_gstr9c
import parse_bo_profile

# parse_ewb_analytics is an optional 4th converter (e-way-bill analytics
# reports). It is not part of the core GSTR-9 / GSTR-9C / BO Profile set,
# so this script must keep working even when that file isn't present on
# a given device/folder.
try:
    import parse_ewb_analytics
    _HAS_EWB = True
except ImportError:
    parse_ewb_analytics = None
    _HAS_EWB = False


def detect_doc_type(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = (pdf.pages[0].extract_text() or '')[:2000]
    if 'GSTR-9C' in text or 'GSTR-9c' in text:
        return 'gstr9c'
    if 'Form GSTR-9' in text or 'Annual Return' in text:
        return 'gstr9'
    if 'OUTWARD SUPPLIES' in text.upper() and 'EWB' in text.upper():
        return 'ewb_analytics'
    if 'Demographic Details' in text or ('GSTIN' in text and 'Overall Risk Score' in text):
        return 'bo_profile'
    # fall back to a second-page scan for reports whose own text may start
    # past page 1 in some exports
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:3]:
            t = page.extract_text() or ''
            if 'Overall Risk Score' in t or 'BIFA' in t:
                return 'bo_profile'
            if 'HSNWISE' in t.upper() and 'EWB' in t.upper():
                return 'ewb_analytics'
    raise ValueError(f"Could not auto-detect document type for {pdf_path}")


CONVERTERS = {
    'gstr9c': parse_gstr9c.convert,
    'gstr9': parse_gstr9.convert,
    'bo_profile': parse_bo_profile.convert,
}
if _HAS_EWB:
    CONVERTERS['ewb_analytics'] = parse_ewb_analytics.convert


def convert_one(pdf_path, out_path):
    doc_type = detect_doc_type(pdf_path)
    if doc_type not in CONVERTERS:
        raise ValueError(
            f"Detected document type '{doc_type}' but its converter "
            f"(parse_{doc_type}.py) is not present next to this script."
        )
    titles = CONVERTERS[doc_type](pdf_path, out_path)
    return doc_type, titles


def process_folder(in_dir, out_dir):
    """Convert every PDF in in_dir, writing .xlsx files into out_dir.
    Used by both the explicit CLI folder-mode and the Pydroid zero-arg mode.
    Returns (converted_count, failed_count).
    """
    os.makedirs(out_dir, exist_ok=True)
    pdfs = sorted(f for f in os.listdir(in_dir) if f.lower().endswith('.pdf'))
    if not pdfs:
        print(f"No PDFs found in {in_dir}")
        return 0, 0

    ok, failed = 0, 0
    for fname in pdfs:
        in_path = os.path.join(in_dir, fname)
        out_path = os.path.join(out_dir, os.path.splitext(fname)[0] + '.xlsx')
        try:
            doc_type, titles = convert_one(in_path, out_path)
            print(f"[{doc_type:12s}] {fname} -> {os.path.basename(out_path)} "
                  f"({len(titles)} sheets)")
            ok += 1
        except Exception as e:
            print(f"[FAILED]      {fname}: {e}")
            traceback.print_exc()
            failed += 1
    return ok, failed


def run_zero_arg_mode():
    """No command-line arguments were given - this is the common case when
    the script is launched by tapping Pydroid's Run button. Use two folders
    next to this script file instead of asking for typed paths."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    in_dir = os.path.join(base_dir, 'PDF_Input')
    out_dir = os.path.join(base_dir, 'Excel_Output')
    os.makedirs(in_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    print("No file/folder argument given - using the default Pydroid folders:")
    print(f"  Input : {in_dir}")
    print(f"  Output: {out_dir}")
    print()

    pdfs = [f for f in os.listdir(in_dir) if f.lower().endswith('.pdf')]
    if not pdfs:
        print("PDF_Input is empty.")
        print("Use Pydroid's file browser (or any file manager) to copy your")
        print("GSTR-9 / GSTR-9C / BO Profile PDF(s) into that folder, then run")
        print("this script again.")
        return

    ok, failed = process_folder(in_dir, out_dir)
    print()
    print(f"Done. {ok} converted, {failed} failed.")
    print(f"Open {out_dir} in your file manager to get the .xlsx files.")


def main():
    if len(sys.argv) < 2:
        run_zero_arg_mode()
        return

    src = sys.argv[1]

    if os.path.isdir(src):
        out_dir = sys.argv[2] if len(sys.argv) > 2 else src
        ok, failed = process_folder(src, out_dir)
        if ok == 0 and failed == 0:
            sys.exit(1)
    else:
        out_path = sys.argv[2] if len(sys.argv) > 2 else (
            os.path.splitext(src)[0] + '.xlsx')
        doc_type, titles = convert_one(src, out_path)
        print(f"Detected: {doc_type}")
        print(f"Wrote {len(titles)} sheets to {out_path}:")
        for t in titles:
            print(f"  - {t}")


if __name__ == '__main__':
    main()
