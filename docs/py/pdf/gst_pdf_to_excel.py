#!/usr/bin/env python3
"""
gst_pdf_to_excel.py - Convert GSTR-9, GSTR-9C, and BO Profile (GST risk
intelligence) PDF reports into clean, properly-sectioned Excel workbooks.

USAGE
    python3 gst_pdf_to_excel.py <input.pdf> [output.xlsx]
    python3 gst_pdf_to_excel.py <folder_of_pdfs> [output_folder]

The document type (GSTR-9C, GSTR-9, or BO Profile) is auto-detected from the
PDF's own text, so this one script works for any future GSTIN/filing you feed
it - no need to touch the code per client.

DESIGN PRINCIPLE
    A worksheet holds the COMPLETE data for exactly one logical heading /
    section of the source PDF. A new heading always starts a new worksheet.
    No section is ever split across two worksheets, and the split points are
    the PDF's own structural markers (item numbers, "Pt." labels, or known
    field/column-header names) - never a fixed row count or page boundary.

Requires: pdfplumber, openpyxl (both already available in this environment).
"""
import os
import sys
import pdfplumber

import parse_gstr9
import parse_gstr9c
import parse_bo_profile
import parse_ewb_analytics


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
    'ewb_analytics': parse_ewb_analytics.convert,
}


def convert_one(pdf_path, out_path):
    doc_type = detect_doc_type(pdf_path)
    titles = CONVERTERS[doc_type](pdf_path, out_path)
    return doc_type, titles


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    src = sys.argv[1]

    if os.path.isdir(src):
        out_dir = sys.argv[2] if len(sys.argv) > 2 else src
        os.makedirs(out_dir, exist_ok=True)
        pdfs = [f for f in os.listdir(src) if f.lower().endswith('.pdf')]
        if not pdfs:
            print(f"No PDFs found in {src}")
            sys.exit(1)
        for fname in pdfs:
            in_path = os.path.join(src, fname)
            out_path = os.path.join(out_dir, os.path.splitext(fname)[0] + '.xlsx')
            try:
                doc_type, titles = convert_one(in_path, out_path)
                print(f"[{doc_type:10s}] {fname} -> {os.path.basename(out_path)} "
                      f"({len(titles)} sheets)")
            except Exception as e:
                print(f"[FAILED]    {fname}: {e}")
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
