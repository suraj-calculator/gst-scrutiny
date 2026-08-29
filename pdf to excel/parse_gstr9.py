import re
import openpyxl
import pdfplumber
from pdf_to_excel_core import extract_flat_rows, extract_verification_rows, write_section, clean_cell

PT_RE = re.compile(r'^Pt\.?[IVXivx]+$')


def _norm(s):
    return re.sub(r'\s+', '', s or '')


ITEM_GROUPS = {
    '4': 'Item 4 - Advances & Outward+Inward Supplies (Tax Payable)',
    '5': 'Item 5 - Outward Supplies (Tax Not Payable)',
    '6': 'Item 6 - ITC Availed During the Year',
    '7': 'Item 7 - ITC Reversed & Ineligible ITC',
    '8': 'Item 8 - Other ITC Related Information',
    '9': 'Item 9 - Tax Paid as Declared in Returns',
    '10': 'Part V - Transactions Declared in Next FY (Items 10-14)',
    '11': 'Part V - Transactions Declared in Next FY (Items 10-14)',
    '12': 'Part V - Transactions Declared in Next FY (Items 10-14)',
    '13': 'Part V - Transactions Declared in Next FY (Items 10-14)',
    '14': 'Part V - Transactions Declared in Next FY (Items 10-14)',
    '15': 'Item 15 - Particulars of Demands and Refunds',
    '16': 'Item 16 - Supplies from Composition Taxpayers etc.',
    '19': 'Item 19 - Late Fee Payable and Paid',
}


def build_sections(rows):
    """The GSTR-9 PDF reprints a "Pt. X" banner plus the column-header rows at
    the top of every page. Sometimes that repeat immediately precedes a brand
    new item (so it IS that item's real header and belongs at the top of its
    sheet); other times it's just a mid-table repeat with more lettered rows
    of the *same* item following (so it's a redundant duplicate, safe to drop).
    `carry` holds rows seen right after a "Pt." banner until we know which case
    we're in.
    """
    sections = []
    current_title = 'Part I - Basic Details (Items 1-3)'
    buffer = []
    carry = []
    after_pt = False

    def flush():
        if buffer:
            sections.append((current_title, list(buffer)))
            buffer.clear()

    for row in rows:
        c0 = _norm(row[0])
        if PT_RE.match(c0):
            after_pt = True
            carry = []
            continue
        if re.match(r'^\d{1,2}$', c0) and c0 in ITEM_GROUPS:
            if ITEM_GROUPS[c0] != current_title:
                flush()
                current_title = ITEM_GROUPS[c0]
                buffer.extend(carry)
            carry = []
            after_pt = False
            buffer.append(row)
            continue
        if after_pt:
            carry.append(row)
            continue
        buffer.append(row)
    flush()
    return sections


def _extract_hsn_note(pdf_path):
    """Items 17 & 18 (HSN-wise summary of outward/inward supplies) carry no
    table in this filing - just a one-line note. Pull it from the page text."""
    lines_out = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            if '17.' in text and 'HSN Wise Summary' in text:
                for raw in text.split('\n'):
                    line = clean_cell(raw.strip())
                    if line and ('HSN Wise Summary' in line or 'download GSTR 9' in line):
                        lines_out.append(line)
    return lines_out


def convert(pdf_path, out_path):
    rows = extract_flat_rows(pdf_path)
    sections = build_sections(rows)

    hsn_note = _extract_hsn_note(pdf_path)
    if hsn_note:
        sections.append(('Items 17-18 - HSN Wise Summary (Note)',
                          [[str(i + 1), t] for i, t in enumerate(hsn_note)]))

    verif = extract_verification_rows(pdf_path)
    if verif:
        sections.append(('Verification', verif))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for title, sec_rows in sections:
        write_section(wb, used, title, None, sec_rows)
    wb.save(out_path)
    return [t for t, _ in sections]


if __name__ == '__main__':
    import sys
    titles = convert(sys.argv[1], sys.argv[2])
    for t in titles:
        print(t)
