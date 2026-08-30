import re
import openpyxl
import pdfplumber
from pdf_to_excel_core import write_section

# The diagonal "EWB ANALYTICS" watermark, extracted with use_text_flow=True,
# always lands as its 12 glyphs together, one per line, spelling the phrase
# forwards - so a run of >=3 of them is unambiguously the watermark.
_MONTH_RE = re.compile(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-?\s*(\d{4})\b')
_NUM_TOKEN_RE = re.compile(r'-?\d+(?:\.\d+)?')

# (number, title-as-printed, shape). Shape drives which stream-consumer runs.
#   'supplies'   - the 3-group x 3-column monthly table (sections 1, 2)
#   'hsn'        - the HSN-wise table (sections 3, 4)
#   'month'      - a plain S.No/Month-Year/N-numeric-columns table
#   'nomonth'    - S.No/N-numeric-columns table with no month column
#   'risk'       - the free-form closing narrative (section 20)
SECTION_HEADERS = [
    ('1', 'OUTWARD SUPPLIES', 'supplies', dict(group3='Exports')),
    ('2', 'INWARD SUPPLIES', 'supplies', dict(group3='Imports')),
    ('3', 'HSNWISE OUTWARD SUPPLIES', 'hsn', {}),
    ('4', 'HSNWISE INWARD SUPPLIES', 'hsn', {}),
    ('5', 'EWB VERIFICATION', 'month', dict(cols=['Total Verification EWB', 'Ok Count', 'Not Ok Count'])),
    ('6', 'Top 50 Suppliers', 'month', dict(cols=[])),
    ('7', 'Top 50 Recipients', 'month', dict(cols=[])),
    ('8', 'CANCELLATIONS & REJECTIONS', 'month', dict(cols=[
        'Cancel EwayBill Count', 'Cancel Assessable Value',
        'Rejection EwayBill Count', 'Rejection Assessable Value'])),
    ('9', 'EWB BLOCKING', 'month', dict(cols=[])),
    ('10', 'EXTENSION OF EWB', 'month', dict(cols=['No. of EWBs', 'Assessable Value'])),
    ('11', 'EXTENSION OF TRANSPORTERS', 'month', dict(cols=[])),
    ('12', 'EWB OF CRTICAL COMMODITIES', 'month', dict(cols=['No. of EWBs', 'Assessable Value', 'Tax Value'])),
    ('13', 'BILL TO SHIP TO TRANSACTIONS', 'month', dict(cols=['No. of EWBs', 'Assessable Value'])),
    ('14', 'OUTWARD B2C TRANSACTIONS', 'month', dict(cols=['No. of EWBs', 'Assessable Value'])),
    ('15', 'INWARD B2C TRANSACTIONS', 'month', dict(cols=[])),
    ('16', 'OUTWARD SUPPLIES TO WORK CONTRACTORS', 'nomonth', dict(cols=['No. of EWBs', 'Assessable Value', 'Tax Value'])),
    ('17', 'ODC EWBs', 'month', dict(cols=[])),
    ('18', 'PART A EWAYBILLS', 'month', dict(cols=['No. of EWBs', 'Assessable Value'])),
    ('19', 'EWB-03 Details', 'month', dict(cols=[])),
    ('20', 'RISK INVOLVED', 'risk', {}),
]


def clean_text(raw):
    text = re.sub(r'(?:^\s*[EWBANLYTICS]\s*$\n?){3,}', '', raw, flags=re.M)
    text = re.sub(r'\bEWB\s*ANALYTICS\b', ' ', text)
    # Strip every known section-title occurrence - they get pre-announced in
    # clusters ahead of their actual data by this PDF's page layout, so they
    # are noise for content-based parsing, not reliable boundaries.
    for num, title, _, _ in SECTION_HEADERS:
        text = re.sub(rf'{re.escape(num)}\.\s*{re.escape(title)}', ' ', text)
    return text


def _numbers(s):
    return _NUM_TOKEN_RE.findall(s)


# ---------------------------------------------------------------------------
# Stream consumers - each takes (text, pos), returns (title_rows, new_pos)
# ---------------------------------------------------------------------------

def consume_supplies(text, pos, group3):
    m = re.search(r'Period\s*2026-2027', text[pos:])
    if not m:
        return [['Month-Year']], pos
    start = pos + m.end()
    total_m = re.search(r'\bTotal\b', text[start:])
    end = start + total_m.start() if total_m else len(text)
    body = text[start:end]

    months = list(_MONTH_RE.finditer(body))
    rows = []
    for i, mm in enumerate(months):
        seg_end = months[i + 1].start() if i + 1 < len(months) else len(body)
        seg = body[mm.end():seg_end]
        toks = _numbers(seg)[:10]
        if len(toks) == 10:
            ewb1, av1a, av1b, tv1, ewb2, av2, tv2, ewb3, av3, tv3 = toks
            av1 = av1a + av1b
        elif len(toks) == 9:
            ewb1, av1, tv1, ewb2, av2, tv2, ewb3, av3, tv3 = toks
        else:
            padded = (toks + ['0'] * 9)[:9]
            ewb1, av1, tv1, ewb2, av2, tv2, ewb3, av3, tv3 = padded
        rows.append([f'{mm.group(1)}-{mm.group(2)}', int(ewb1), int(av1), int(tv1),
                     int(ewb2), int(av2), int(tv2), int(ewb3), int(av3), int(tv3)])

    header = ['Month-Year',
              'Supplies - No. of EWBs', 'Supplies - Assessable Value', 'Supplies - Tax Value',
              'Non-Supplies - No. of EWBs', 'Non-Supplies - Assessable Value', 'Non-Supplies - Tax Value',
              f'{group3} - No. of EWBs', f'{group3} - Assessable Value', f'{group3} - Tax Value']
    total = ['Total'] + [sum(r[i] for r in rows) for i in range(1, 10)]
    new_pos = start + (total_m.end() if total_m else 0)
    if total_m:
        rest = text[new_pos:new_pos + 200]
        skip = re.match(r'[\d\s]*', rest)
        new_pos += skip.end() if skip else 0
    return [header] + rows + [total], new_pos


def consume_hsn(text, pos):
    m = re.search(r'S\.No\.\s*HSN Code\s*HSN Description', text[pos:])
    if not m:
        return [['S.No', 'HSN Code', 'HSN Description', 'EWB Count', 'Assessable Value', 'Tax Value']], pos
    start = pos + m.end()

    # Search for each expected S.No in strict sequence, one at a time. A
    # single shared regex pass is ambiguous here: a previous entry's own
    # trailing digit (e.g. the "0" ending "...400020 0") can accidentally
    # look like "<srno>\n<hsn code>" and get consumed before the real anchor
    # is ever reached. Searching for the exact expected number removes that
    # ambiguity.
    filtered = []
    cursor = start
    expected = 1
    while True:
        anchor_re = re.compile(rf'(?<!\d){expected}\s*\n\s*(\d{{2,8}})\b')
        window_end = cursor + 3000  # one HSN entry is long-winded but bounded
        mm = anchor_re.search(text, cursor, window_end)
        if not mm:
            break
        filtered.append((mm.start(), mm.end(), str(expected), mm.group(1)))
        cursor = mm.end()
        expected += 1
        if expected > 100:  # safety valve
            break
    if not filtered:
        return [['S.No', 'HSN Code', 'HSN Description', 'EWB Count', 'Assessable Value', 'Tax Value']], start

    entries_end = filtered[-1][1]
    total_m = re.search(r'Total\s+(\d+)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)', text[entries_end:entries_end + 4000])

    rows = []
    for i, (a_start, a_end, srno, hsn) in enumerate(filtered):
        seg_end = filtered[i + 1][0] if i + 1 < len(filtered) else (
            entries_end + (total_m.start() if total_m else 0))
        seg = text[a_end:seg_end]
        num_matches = list(_NUM_TOKEN_RE.finditer(seg))
        if len(num_matches) >= 3:
            last3 = num_matches[-3:]
            ewb_count, av, tv = (m.group() for m in last3)
            desc_text = seg[:last3[0].start()]
        else:
            ewb_count, av, tv = '0', '0', '0'
            desc_text = seg
        desc = re.sub(r'\s+', ' ', desc_text).strip(' -\n')
        rows.append([int(srno), hsn, desc, int(ewb_count), int(av), int(tv)])

    header = ['S.No', 'HSN Code', 'HSN Description', 'EWB Count', 'Assessable Value', 'Tax Value']
    total = ['', '', 'Total', sum(r[3] for r in rows), sum(r[4] for r in rows), sum(r[5] for r in rows)]
    new_pos = entries_end + (total_m.end() if total_m else 0)
    return [header] + rows + [total], new_pos


def consume_month(text, pos, cols):
    """A plain Month-Year table with len(cols) numeric columns, OR (if the
    section has no data) a 'Details Not Found !!' marker - whichever comes
    first from `pos` is what this section actually is."""
    nf_m = re.search(r'Details Not Found\s*!*', text[pos:])
    month_m = _MONTH_RE.search(text[pos:])
    nf_pos = pos + nf_m.start() if nf_m else None
    month_pos = pos + month_m.start() if month_m else None

    if nf_pos is not None and (month_pos is None or nf_pos < month_pos):
        header = ['Month-Year'] + cols
        new_pos = pos + nf_m.end()
        return [header, ['Details Not Found'] + [''] * len(cols)], new_pos

    if month_pos is None:
        return [['Month-Year'] + cols], pos

    n_cols = len(cols)
    start = month_pos
    total_m = re.search(r'\bTotal\b', text[start:])
    end = start + total_m.start() if total_m else len(text)
    body = text[start:end]
    months = list(_MONTH_RE.finditer(body))
    rows = []
    for i, mm in enumerate(months):
        seg_end = months[i + 1].start() if i + 1 < len(months) else len(body)
        seg = body[mm.end():seg_end]
        toks = _numbers(seg)[:n_cols]
        toks = (toks + ['0'] * n_cols)[:n_cols]
        rows.append([f'{mm.group(1)}-{mm.group(2)}'] + [int(t) for t in toks])
    header = ['Month-Year'] + cols
    total = ['Total'] + [sum(r[i] for r in rows) for i in range(1, n_cols + 1)]
    new_pos = start + (total_m.end() if total_m else 0)
    if total_m:
        rest = text[new_pos:new_pos + 200]
        skip = re.match(r'[\d\s.]*', rest)
        new_pos += skip.end() if skip else 0
    return [header] + rows + [total], new_pos


def consume_nomonth(text, pos, cols):
    nf_m = re.search(r'Details Not Found\s*!*', text[pos:])
    header_m = re.search(r'S\.No\.\s*No\. of EWBs', text[pos:])
    nf_pos = pos + nf_m.start() if nf_m else None
    h_pos = pos + header_m.start() if header_m else None

    if nf_pos is not None and (h_pos is None or nf_pos < h_pos):
        header = ['S.No'] + cols
        new_pos = pos + nf_m.end()
        return [header, ['Details Not Found'] + [''] * len(cols)], new_pos

    if header_m is None:
        return [['S.No'] + cols], pos

    start = pos + header_m.end()
    total_m = re.search(r'\bTotal\b', text[start:])
    end = start + total_m.start() if total_m else len(text)
    body = text[start:end]
    n_cols = len(cols)
    toks = _numbers(body)
    rows = []
    i = 0
    while i + n_cols < len(toks):
        srno = toks[i]
        vals = toks[i + 1: i + 1 + n_cols]
        rows.append([int(srno)] + [int(v) for v in vals])
        i += 1 + n_cols
    header = ['S.No'] + cols
    total = ['Total'] + [sum(r[i] for r in rows) for i in range(1, n_cols + 1)]
    new_pos = start + (total_m.end() if total_m else 0)
    return [header] + rows + [total], new_pos


def consume_risk(text, pos):
    rows = [['Category', 'Detail']]
    tail = text[pos:]
    m = re.search(r'Abnormal growth.*?\(> ?75% of the previous month ?\)', tail)
    if m:
        growth_text = tail[m.end():]
        cutoff = growth_text.find('Sales to URP')
        if cutoff != -1:
            growth_text = growth_text[:cutoff]
        # The PDF's flow order prints "<year> <month>" as the label for the
        # NEXT count/value pair, attached to the tail of the CURRENT line -
        # so labels and data pairs are offset by one and must be re-paired
        # by position, not by proximity in the raw text.
        labels = re.findall(r'(\d{4})\s+([A-Za-z]+)', growth_text)
        pairs = re.findall(r'(\d+)\s+(\d+\.\d{2})', growth_text)
        for (year, month), (count, value) in zip(labels, pairs):
            rows.append([f'{month} {year} - Abnormal EWB growth (>75% of previous month)',
                         f'EWBs: {count}, Assessable Value: Rs. {value}'])
    if 'Sales to URP' in tail:
        after = tail[tail.find('Sales to URP'):]
        rows.append(['Sales to URP (Individual invoice value > Rs 10 Lakh)',
                      'No entries reported for this period'])
    return rows, len(text)


# ---------------------------------------------------------------------------
# Top-level conversion
# ---------------------------------------------------------------------------

def convert(pdf_path, out_path):
    with pdfplumber.open(pdf_path) as pdf:
        raw_text = '\n'.join((p.extract_text(use_text_flow=True) or '') for p in pdf.pages)
    text = clean_text(raw_text)

    pos = 0
    sections = []
    for num, title_raw, shape, kw in SECTION_HEADERS:
        title = f'{num}. {title_raw}'
        if shape == 'supplies':
            rows, pos = consume_supplies(text, pos, kw['group3'])
        elif shape == 'hsn':
            rows, pos = consume_hsn(text, pos)
        elif shape == 'month':
            rows, pos = consume_month(text, pos, kw['cols'])
        elif shape == 'nomonth':
            rows, pos = consume_nomonth(text, pos, kw['cols'])
        elif shape == 'risk':
            rows, pos = consume_risk(text, pos)
        sections.append((title, rows))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for title, rows in sections:
        str_rows = [[('' if c is None else str(c)) for c in row] for row in rows]
        write_section(wb, used, title, None, str_rows)
    wb.save(out_path)
    return [t for t, _ in sections]


if __name__ == '__main__':
    import sys
    titles = convert(sys.argv[1], sys.argv[2])
    for t in titles:
        print(t)
