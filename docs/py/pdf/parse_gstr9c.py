import re
import openpyxl
from pdf_to_excel_core import extract_flat_rows, extract_verification_rows, write_section

PT_RE = re.compile(r'^Pt\.?[IVXivx]+$')


def _norm(s):
    return re.sub(r'\s+', '', s or '')


# Item number (as it appears in column 0) -> canonical worksheet title.
# Items not listed here (1, 2, 3(a-d), 4) stay grouped under Part I.
ITEM_GROUPS = {
    '5': 'Item 5 - Reconciliation of Gross Turnover',
    '6': 'Item 6 - Reasons for Turnover Difference',
    '7': 'Item 7 - Reconciliation of Taxable Turnover',
    '8': 'Item 8 - Reasons for Taxable Turnover Difference',
    '9': 'Item 9 - Rate-wise Tax Liability & Amount Payable',
    '10': 'Item 10 - Reasons for Un-reconciled Payment',
    '11': 'Item 11 - Additional Amount Payable but Not Paid',
    '12': 'Item 12 - Reconciliation of Net ITC',
    '13': 'Item 13 - Reasons for Un-reconciled ITC Difference',
    '14': 'Item 14 - ITC Reconciliation with Books/Expenses',
    '15': 'Item 15 - Reasons for Un-reconciled ITC (Expense Basis)',
    '16': 'Item 16 - Tax Payable on Un-reconciled ITC Difference',
    '17': 'Item 17 - Late Fee Payable and Paid',
}
HEADLESS_PT = {
    'Pt.V': 'Part V - Additional Liability due to Non-reconciliation',
}


def build_sections(rows):
    sections = []
    current_title = 'Part I - Basic Details (Items 1-4)'
    buffer = []

    def flush():
        if buffer:
            sections.append((current_title, list(buffer)))
            buffer.clear()

    for row in rows:
        c0 = _norm(row[0])
        if PT_RE.match(c0):
            if c0 in HEADLESS_PT:
                flush()
                current_title = HEADLESS_PT[c0]
            continue  # Pt banner row itself carries no data - drop it
        if re.match(r'^\d{1,2}$', c0) and c0 in ITEM_GROUPS:
            if ITEM_GROUPS[c0] != current_title:
                flush()
                current_title = ITEM_GROUPS[c0]
            buffer.append(row)
            continue
        buffer.append(row)
    flush()
    return sections


def convert(pdf_path, out_path):
    rows = extract_flat_rows(pdf_path)
    sections = build_sections(rows)

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
