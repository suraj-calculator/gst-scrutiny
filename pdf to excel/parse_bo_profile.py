import re
import openpyxl
from pdf_to_excel_core import extract_flat_rows, write_section


def _cell_has(row, *needles):
    for c in row:
        if not c:
            continue
        for n in needles:
            if n.lower() in c.lower():
                return True
    return False


def _row0_eq(row, *values):
    c0 = (row[0] or '').strip() if row and row[0] else ''
    return c0 in values


# Ordered boundary markers. Each tuple:
#   (canonical_title, matcher(row)->bool, keep_trigger_row, note_col)
# `keep_trigger_row`=True means the row that triggered the boundary is itself
# real data/header and stays as the section's first row; False means it is a
# pure title banner (optionally with a note in note_col) that gets dropped
# from the data and used only to label the sheet.
BOUNDARIES = [
    ('Flagged in Lead Based Dashboard', lambda r: _row0_eq(r, 'Lead Type'), True, None),
    ('Pre-GST Details', lambda r: _row0_eq(r, 'Central Excise Reg Number'), True, None),
    ('Places of Business', lambda r: _row0_eq(r, 'Principal Place of Business'), True, None),
    ('Bank Account Details', lambda r: _row0_eq(r, 'Bank Name'), True, None),
    ('Member Details', lambda r: _row0_eq(r, 'Member Name'), True, None),
    ('Shared Entity', lambda r: _row0_eq(r, 'Shared Entity'), True, None),
    ('Shared Members', lambda r: _row0_eq(r, 'GSTIN') and len(r) > 1 and r[1] == 'Other GSTIN', True, None),
    ('Financial Information', lambda r: _row0_eq(r, 'Financial Information'), False, 1),
    ('BIFA Specific Information', lambda r: r[0] and r[0].startswith('BIFA'), False, 1),
    ('ITC Passed On', lambda r: _row0_eq(r, 'ITC Passed On'), False, 1),
    ('ITC Received', lambda r: _row0_eq(r, 'ITC Received'), False, 1),
    ('Top 10 Beneficiaries based on ITC Passed (last 12 months)',
     lambda r: r[0] and r[0].startswith(('Top 10 Benef', 'Top 10 Benefciaries')), False, 1),
    ('Top 10 Suppliers based on ITC Received (last 12 months)',
     lambda r: r[0] and r[0].startswith('Top 10 Suppliers'), False, 1),
    ('EWB Related Information', lambda r: _cell_has(r, 'EWB Related Information'), False, None),
    ('E-Invoice Related Information', lambda r: _cell_has(r, 'E-Invoice Related Information'), False, None),
    ('Refund Details', lambda r: _cell_has(r, 'Refund Details'), False, None),
    ('ITC Received from Related / Cancelled Party', lambda r: _cell_has(r, 'ITC Received from Related'), False, None),
    ('ITC Passed On to Related / Cancelled Party', lambda r: _cell_has(r, 'ITC Passed On to Related'), False, None),
    ('HSN as per REG01', lambda r: list(filter(None, r)) == ['HSN', 'Description'], True, None),
    ('Top 10 HSN as per GSTR-1 (last 12 months)', lambda r: _cell_has(r, 'Top 10 HSN as per GSTR-1'), False, None),
    ('Top 10 HSN as per EWB (last 12 months)', lambda r: _cell_has(r, 'Top 10 HSN as per EWB'), False, None),
    ('Top 10 HSN as per E-Invoice (last 12 months)', lambda r: _cell_has(r, 'Top 10 HSN as per E-Invoice'), False, None),
    ('Export / ICEGATE Information', lambda r: _cell_has(r, 'Export / ICEGATE', 'Export/ICEGATE'), False, None),
    # "Appeal Information" is floating page text (no table border of its own in
    # this filing), so detect it via its distinctive column-header signature
    # instead of a title row.
    ('Appeal Information', lambda r: _row0_eq(r, 'Financial Year') and len(r) > 1 and r[1] == 'ARN', True, None),
    ('Case Information', lambda r: _row0_eq(r, 'Case Information'), False, 1),
    ('DRC Payment Information', lambda r: _cell_has(r, 'DRC Payment Information'), False, None),
    # "Transfer Information" is likewise floating text - detect via its header row.
    ('Transfer Information', lambda r: _row0_eq(r, 'Recommend Date'), True, None),
]

# "Return Filing Details" has two identical-looking sub-tables (GSTR-3B, then
# GSTR-1), each introduced by floating text that never lands in any table row.
# Both share the exact same header row, so they can only be told apart by
# which occurrence it is.
_RETURN_HEADER = ['Date of Filing', 'Return Period', 'IP Address']


def build_sections(rows):
    sections = []
    current_title = 'Demographic Details'
    buffer = []
    b_idx = 0
    return_header_seen = 0

    def flush():
        if buffer:
            sections.append((current_title, list(buffer)))
            buffer.clear()

    for row in rows:
        padded_for_cmp = [c for c in row[:3]]
        if padded_for_cmp == _RETURN_HEADER:
            return_header_seen += 1
            flush()
            current_title = ('Return Filing Details - GSTR-3B' if return_header_seen == 1
                              else 'Return Filing Details - GSTR-1')
            buffer.append(row)
            continue
        if b_idx < len(BOUNDARIES) and BOUNDARIES[b_idx][1](row):
            title, _, keep, note_idx = BOUNDARIES[b_idx]
            flush()
            current_title = title
            b_idx += 1
            if keep:
                buffer.append(row)
            continue
        buffer.append(row)
    flush()
    return sections


def convert(pdf_path, out_path):
    rows = extract_flat_rows(pdf_path, drop_letterhead=True)
    sections = build_sections(rows)

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
