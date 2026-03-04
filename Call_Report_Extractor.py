# Call_Report_Extractor.py
import io
import os
import re
import zipfile
import traceback
from typing import Dict, List, Tuple, Optional

import streamlit as st
import pandas as pd
import pdfplumber
from bs4 import BeautifulSoup
import html as html_std  # for html.escape


# ==============================
# STREAMLIT CONFIG
# ==============================
st.set_page_config(
    page_title="PDF → HTML → Excel (5300 Extractor)",
    page_icon="📄",
    layout="wide",
)
st.title("Call Report Data Extractor")
st.caption("Select PDFs → Convert to HTML (via pdfplumber) → Extract account codes → Download Excel")


# ==============================
# ACCOUNT CODE LIST (as provided)
# ==============================
account_codes = ['002', '003', '007', '008', '009A', '009B', '009C', '009D2', '010', '010A', '010B', '010C', '013',
                 '013A', '013B1', '013B2', '014', '018', '018A', '018B1', '018B2', '020B', '020C1', '020C2', '020D',
                 '020T', '021B', '021C1', '021C2', '021D', '021T', '022B', '022C1', '022C2', '022D', '022T', '023B',
                 '023C1', '023C2', '023D', '023T', '024B', '025A', '025A1', '025B', '025B1', '026B', '027B', '028B',
                 '031A', '031B', '031C', '031D', '034E', '035E1', '035E2', '041A', '041B', '041C1', '041C2', '041D',
                 '041E', '041T', '042A5', '042A6', '042A7', '042A8', '042A9', '045A', '045B', '053E', '058A', '058B1',
                 '058B2', '058C', '065A4', '067A2', '068A', '069A', '083', '084', '089B', '090A1', '090B1', '090C5',
                 '090C6', '090C7', '090C8', '090H2', '090H3', '090J2', '090J3', '090K2', '090K3', '090L2', '090L3',
                 '090M', '090M1', '099A5', '099A6', '099A7', '099A8', '1000F', '1001F', '1004', '1004A', '1004B',
                 '1004C', '1030', '1030C', '1050', '1050A', '1051', '1052', '1053', '1054', '1055', '1056', '1061',
                 '1061A', '1062', '1062A', '1063', '1063A', '110', '115', '117', '119', '120', '127B', '128B', '129B',
                 '130A', '130B', '131', '136', '137', '143A3', '143A4', '143B3', '143B4', '143C3', '143C4', '143D3',
                 '143D4', '210', '230', '250', '260', '270', '280', '290', '300', '310', '320', '340', '350', '360',
                 '370', '380', '381', '385', '386A', '386B', '396', '397', '397A', '400A', '400A1', '400B1', '400C5',
                 '400C6', '400C7', '400C8', '400H2', '400H3', '400J2', '400J3', '400L2', '400L3', '400M', '400M1',
                 '400P', '400T1', '420A', '420B', '420C', '421', '430', '431', '440', '451', '452', '453', '454', '455',
                 '457', '458', '460', '463A5', '463A6', '463A7', '463A8', '475A1', '475B1', '475C5', '475C6', '475C7',
                 '475C8', '475H2', '475H3', '475J2', '475J3', '475K2', '475K3', '475L2', '475L3', '475M', '475M1', '521',
                 '522', '522A', '523', '524', '525', '526', '550', '550C1', '550C2', '550D', '550E', '550F', '550T',
                 '551', '551C1', '551C2', '551D', '551E', '551F', '551T', '562A', '562B', '563A', '564A', '564B', '565',
                 '566B', '567', '568', '595A', '595B', '602', '617A', '618A', '630', '630A', '630B1', '630B2', '631',
                 '632', '636', '638', '639', '641', '643', '644', '657', '657A', '658', '658A', '661A', '668', '671',
                 '680', '681', '690', '691', '691B1', '691C1', '691C2', '691L', '691L2', '691L7', '691L8', '691L9',
                 '691N', '691N2', '691N7', '691N8', '691N9', '691P', '691P1', '691P2', '698A', '698C', '700', '701',
                 '703A', '704A2', '704C1', '704C2', '704C3', '718A3', '718A4', '718A5', '719', '730A', '730B', '730B1',
                 '730B2', '736', '769A', '769B', '779', '779A', '784A', '788', '789C', '789D', '789E', '789E1', '789E2',
                 '789F', '789G', '789H', '794', '798A', '801', '811D', '812C', '814K', '815C', '816A', '816B5', '816T',
                 '818A', '819', '820A', '822C', '825', '851', '852', '853', '860A', '860B1', '860B2', '860C', '865A',
                 '867A', '867B1', '867B2', '867C', '875', '876', '877', '878', '880', '880A', '880B1', '880B2', '881',
                 '883A', '883B1', '883B2', '883C', '884', '884C', '884D', '885A', '885A1', '885A2', '885A3', '885A4',
                 '900A1', '900B1', '900C5', '900C6', '900C7', '900C8', '900H2', '900H3', '900J2', '900J3', '900K2',
                 '900K3', '900K4', '900L2', '900L3', '900M', '900M1', '900P', '900T1', '902', '902A', '906A', '906B1',
                 '906B2', '906C', '908A', '908B1', '908B2', '908C', '911', '911A', '925A', '926', '927', '928', '940',
                 '945A', '945B', '945C', '954', '956', '958', '959A', '960A', '960B', '961A5', '961A6', '961A7',
                 '961A8', '961A9', '963A', '963C', '966', '968', '971', '993', '994', '994A', '995', '996', '997',
                 '998', 'AS0003', 'AS0004', 'AS0005', 'AS0007', 'AS0008', 'AS0009', 'AS0010', 'AS0013', 'AS0016',
                 'AS0017', 'AS0022', 'AS0023', 'AS0024', 'AS0025', 'AS0032', 'AS0036', 'AS0041', 'AS0042', 'AS0048',
                 'AS0050', 'AS0051', 'AS0052', 'AS0053', 'AS0054', 'AS0055', 'AS0056', 'AS0057', 'AS0058', 'AS0059',
                 'AS0060', 'AS0061', 'AS0062', 'AS0063', 'AS0064', 'AS0065', 'AS0066', 'AS0067', 'AS0068', 'AS0069',
                 'AS0070', 'AS0071', 'AS0072', 'AS0073', 'BA0009', 'CH0007', 'CH0008', 'CH0015', 'CH0016', 'CH0017',
                 'CH0018', 'CH0019', 'CH0020', 'CH0021', 'CH0022', 'CH0023', 'CH0024', 'CH0025', 'CH0026', 'CH0027',
                 'CH0028', 'CH0029', 'CH0030', 'CH0031', 'CH0032', 'CH0033', 'CH0034', 'CH0035', 'CH0036', 'CH0037',
                 'CH0038', 'CH0039', 'CH0040', 'CH0047', 'CH0048', 'CM0099', 'DL0002', 'DL0009', 'DL0016', 'DL0022',
                 'DL0023', 'DL0024', 'DL0025', 'DL0026', 'DL0027', 'DL0028', 'DL0030', 'DL0037', 'DL0044', 'DL0050',
                 'DL0051', 'DL0052', 'DL0053', 'DL0054', 'DL0055', 'DL0056', 'DL0057', 'DL0058', 'DL0059', 'DL0060',
                 'DL0061', 'DL0062', 'DL0063', 'DL0064', 'DL0065', 'DL0066', 'DL0067', 'DL0068', 'DL0069', 'DL0070',
                 'DL0071', 'DL0072', 'DL0073', 'DL0074', 'DL0075', 'DL0076', 'DL0077', 'DL0078', 'DL0079', 'DL0080',
                 'DL0081', 'DL0082', 'DL0083', 'DL0084', 'DL0085', 'DL0086', 'DL0087', 'DL0088', 'DL0089', 'DL0090',
                 'DL0091', 'DL0092', 'DL0093', 'DL0094', 'DL0095', 'DL0096', 'DL0097', 'DL0098', 'DL0099', 'DL0100',
                 'DL0101', 'DL0102', 'DL0103', 'DL0104', 'DL0105', 'DL0106', 'DL0107', 'DL0108', 'DL0109', 'DL0110',
                 'DL0111', 'DL0112', 'DL0113', 'DL0114', 'DL0115', 'DL0116', 'DL0117', 'DL0118', 'DL0119', 'DL0120',
                 'DL0122', 'DL0123', 'DL0124', 'DL0125', 'DL0126', 'DL0127', 'DL0128', 'DL0129', 'DL0130', 'DL0131',
                 'DL0132', 'DL0133', 'DL0134', 'DL0135', 'DL0136', 'DL0137', 'DL0138', 'DL0139', 'DL0140', 'DL0141',
                 'DL0142', 'DL0144', 'DL0145', 'DL0146', 'DL0147', 'DL0148', 'DL0149', 'DT0001', 'DT0002', 'DT0003',
                 'DT0004', 'DT0005', 'DT0006', 'DT0007', 'DT0008', 'DT0009', 'DT0010', 'DT0011', 'DT0012', 'DT0013',
                 'DT0014', 'DT0015', 'DT0016', 'EQ0009', 'IN0001', 'IN0002', 'IN0003', 'IN0004', 'IN0005', 'IN0006',
                 'IN0007', 'IN0008', 'IS0005', 'IS0010', 'IS0011', 'IS0012', 'IS0013', 'IS0016', 'IS0017', 'IS0020',
                 'IS0029', 'IS0030', 'IS0046', 'IS0047', 'IS0048', 'IS0049', 'LC0047', 'LC0085', 'LI0003', 'LI0069',
                 'LN0050', 'LN0051', 'LN0052', 'LN0053', 'LN0054', 'LN0055', 'LN0056', 'LN0057', 'LQ0013', 'LQ0014',
                 'LQ0015', 'LQ0016', 'LQ0017', 'LQ0018', 'LQ0019', 'LQ0020', 'LQ0021', 'LQ0022', 'LQ0023', 'LQ0024',
                 'LQ0025', 'LQ0026', 'LQ0027', 'LQ0028', 'LQ0029', 'LQ0030', 'LQ0035', 'LQ0039', 'LQ0040', 'LQ0043',
                 'LQ0044', 'LQ0045', 'LQ0046', 'LQ0047', 'LQ0053', 'LQ0060', 'LQ0061', 'LQ0062', 'LQ0860', 'LR0001',
                 'LR0002', 'LR0003', 'LR0004', 'LR0005', 'LR0006', 'LR0007', 'LR0008', 'NV0001', 'NV0002', 'NV0003',
                 'NV0004', 'NV0013', 'NV0014', 'NV0015', 'NV0016', 'NV0017', 'NV0018', 'NV0019', 'NV0020', 'NV0021',
                 'NV0022', 'NV0023', 'NV0024', 'NV0025', 'NV0026', 'NV0027', 'NV0028', 'NV0029', 'NV0030', 'NV0031',
                 'NV0032', 'NV0033', 'NV0034', 'NV0035', 'NV0036', 'NV0037', 'NV0038', 'NV0039', 'NV0040', 'NV0041',
                 'NV0042', 'NV0043', 'NV0044', 'NV0045', 'NV0046', 'NV0047', 'NV0048', 'NV0049', 'NV0050', 'NV0051',
                 'NV0052', 'NV0053', 'NV0054', 'NV0055', 'NV0056', 'NV0057', 'NV0058', 'NV0059', 'NV0060', 'NV0061',
                 'NV0062', 'NV0063', 'NV0064', 'NV0065', 'NV0066', 'NV0067', 'NV0068', 'NV0069', 'NV0070', 'NV0071',
                 'NV0072', 'NV0073', 'NV0074', 'NV0075', 'NV0076', 'NV0077', 'NV0078', 'NV0079', 'NV0080', 'NV0081',
                 'NV0083', 'NV0084', 'NV0087', 'NV0088', 'NV0089', 'NV0090', 'NV0091', 'NV0092', 'NV0093', 'NV0094',
                 'NV0095', 'NV0096', 'NV0097', 'NV0098', 'NV0099', 'NV0100', 'NV0101', 'NV0102', 'NV0103', 'NV0104',
                 'NV0105', 'NV0106', 'NV0107', 'NV0108', 'NV0109', 'NV0110', 'NV0111', 'NV0112', 'NV0113', 'NV0114',
                 'NV0115', 'NV0116', 'NV0122', 'NV0128', 'NV0134', 'NV0140', 'NV0141', 'NV0142', 'NV0143', 'NV0144',
                 'NV0145', 'NV0146', 'NV0153', 'NV0154', 'NV0155', 'NV0156', 'NV0157', 'NV0158', 'NV0159', 'NV0160',
                 'NV0161', 'NV0162', 'NV0169', 'NV0170', 'NV0172', 'NV0173', 'NW0001', 'NW0002', 'NW0004', 'NW0010',
                 'PC0001', 'PC0002', 'PC0003', 'PC0004', 'PC0005', 'PC0006', 'PC0007', 'PC0008', 'PC0009', 'PC0010',
                 'RB0001', 'RB0002', 'RB0003', 'RB0004', 'RB0005', 'RB0006', 'RB0007', 'RB0008', 'RB0009', 'RB0010',
                 'RB0011', 'RB0012', 'RB0013', 'RB0014', 'RB0015', 'RB0016', 'RB0017', 'RB0018', 'RB0019', 'RB0020',
                 'RB0021', 'RB0022', 'RB0023', 'RB0024', 'RB0025', 'RB0026', 'RB0027', 'RB0028', 'RB0029', 'RB0030',
                 'RB0031', 'RB0032', 'RB0033', 'RB0034', 'RB0035', 'RB0036', 'RB0037', 'RB0038', 'RB0039', 'RB0040',
                 'RB0041', 'RB0042', 'RB0043', 'RB0044', 'RB0045', 'RB0046', 'RB0047', 'RB0048', 'RB0049', 'RB0050',
                 'RB0051', 'RB0052', 'RB0053', 'RB0054', 'RB0055', 'RB0056', 'RB0057', 'RB0058', 'RB0059', 'RB0060',
                 'RB0061', 'RB0062', 'RB0063', 'RB0064', 'RB0065', 'RB0066', 'RB0067', 'RB0068', 'RB0069', 'RB0070',
                 'RB0071', 'RB0072', 'RB0073', 'RB0074', 'RB0075', 'RB0076', 'RB0077', 'RB0078', 'RB0079', 'RB0080',
                 'RB0081', 'RB0082', 'RB0083', 'RB0084', 'RB0085', 'RB0086', 'RB0087', 'RB0088', 'RB0089', 'RB0090',
                 'RB0091', 'RB0092', 'RB0093', 'RB0094', 'RB0095', 'RB0096', 'RB0097', 'RB0098', 'RB0099', 'RB0100',
                 'RB0101', 'RB0102', 'RB0103', 'RB0104', 'RB0105', 'RB0106', 'RB0107', 'RB0108', 'RB0109', 'RB0110',
                 'RB0111', 'RB0112', 'RB0113', 'RB0114', 'RB0115', 'RB0116', 'RB0117', 'RB0118', 'RB0119', 'RB0120',
                 'RB0121', 'RB0122', 'RB0123', 'RB0124', 'RB0125', 'RB0126', 'RB0127', 'RB0128', 'RB0129', 'RB0130',
                 'RB0131', 'RB0132', 'RB0133', 'RB0134', 'RB0135', 'RB0136', 'RB0137', 'RB0138', 'RB0139', 'RB0140',
                 'RB0141', 'RB0142', 'RB0143', 'RB0144', 'RB0145', 'RB0146', 'RB0147', 'RB0148', 'RB0149', 'RB0150',
                 'RB0151', 'RB0152', 'RB0153', 'RB0154', 'RB0155', 'RB0156', 'RB0157', 'RB0158', 'RB0159', 'RB0160',
                 'RB0161', 'RB0162', 'RB0163', 'RB0164', 'RB0165', 'RB0166', 'RB0167', 'RB0168', 'RB0169', 'RB0170',
                 'RB0171', 'RB0172', 'RB0177', 'RL0001', 'RL0002', 'RL0003', 'RL0004', 'RL0005', 'RL0006', 'RL0007',
                 'RL0008', 'RL0009', 'RL0010', 'RL0011', 'RL0012', 'RL0013', 'RL0014', 'RL0015', 'RL0016', 'RL0017',
                 'RL0018', 'RL0019', 'RL0020', 'RL0021', 'RL0022', 'RL0023', 'RL0024', 'RL0025', 'RL0026', 'RL0027',
                 'RL0028', 'RL0029', 'RL0030', 'RL0031', 'RL0032', 'RL0033', 'RL0034', 'RL0035', 'RL0036', 'RL0037',
                 'RL0038', 'RL0039', 'RL0040', 'RL0041', 'RL0042', 'RL0043', 'RL0044', 'RL0045', 'RL0046', 'RL0047',
                 'RL0048', 'RL0050', 'SH0013', 'SH0018', 'SH0880', 'SL0012', 'SL0013', 'SL0014', 'SL0015', 'SL0018',
                 'SL0019', 'SL0020', 'SL0021', 'SL0022', 'SL0023', 'SL0024', 'SL0026', 'SL0028', 'SL0029', 'SL0030',
                 'SL0032', 'SL0033', 'SL0034', 'SL0035', 'SL0036', 'SL0037', 'SL0038', 'SL0039', 'SL0041', 'SL0043',
                 'SL0045', 'SL0047', 'SL0049', 'SL0051', 'SL0053', 'SL0055', 'SL0056', 'SL0057', 'SL0058', 'SL0059']


# ==============================
# HELPERS
# ==============================
def sanitize_sheet_name(name: str) -> str:
    """Valid Excel sheet name: remove invalid chars and limit to 31 chars."""
    invalid = r'[]:*?/\\'
    for ch in invalid:
        name = name.replace(ch, '_')
    name = name.strip().strip("'")
    return name[:31] if len(name) > 31 else name

def make_unique_sheet_name(base: str, used: set) -> str:
    """Ensure sheet name unique within workbook, respecting 31-char limit."""
    name = sanitize_sheet_name(base)
    if name not in used:
        used.add(name)
        return name
    counter = 2
    while True:
        suffix = f"_{counter}"
        max_base_len = 31 - len(suffix)
        candidate = sanitize_sheet_name(name[:max_base_len] + suffix)
        if candidate not in used:
            used.add(candidate)
            return candidate
        counter += 1

def parse_numeric(text: str):
    """
    Convert currency/number strings to float if possible.
    Handles $, commas, and parentheses negatives. Returns float or original string.
    """
    if text is None:
        return "null"
    cleaned = text.strip()
    negative = cleaned.startswith('(') and cleaned.endswith(')')
    cleaned = cleaned.replace('(', '').replace(')', '')
    cleaned = cleaned.replace('$', '').replace(',', '').strip()
    if re.fullmatch(r'[-+]?\d*\.?\d+', cleaned):
        val = float(cleaned) if cleaned != '' else 0.0
        return -val if negative else val
    return text


# ==============================
# PDF → HTML (via pdfplumber), returning HTML string
# ==============================
def pdf_to_html_string(pdf_bytes: bytes) -> str:
    """
    Convert a single PDF (bytes) to an HTML string using pdfplumber for text and tables.
    All content is escaped; HTML uses real tags (<p>, <table>, <br>).
    """
    parts = [
        "<!DOCTYPE html>",
        "<html><head><meta charset='utf-8'><title>Converted PDF</title></head><body>"
    ]
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                safe_text = html_std.escape(text).replace("\n", "<br>")
                parts.append(f"<p>{safe_text}</p>")

            tables = page.extract_tables()
            for table in (tables or []):
                parts.append("<table border='1' cellspacing='0' cellpadding='4'>")
                for row in table:
                    parts.append("<tr>")
                    for cell in (row or []):
                        safe_cell = "" if cell is None else html_std.escape(str(cell))
                        parts.append(f"<td>{safe_cell}</td>")
                    parts.append("</tr>")
                parts.append("</table>")
    parts.append("</body></html>")
    return "".join(parts)

@st.cache_data(show_spinner=False)
def cached_pdf_to_html(name: str, file_bytes: bytes) -> str:
    # Cache by file name + bytes so re-renders don't redo the work
    return pdf_to_html_string(file_bytes)


# ==============================
# HTML → rows (extract account/value pairs)
# ==============================
def extract_rows_from_html(html: str) -> List[Tuple[str, object]]:
    """
    From HTML, find any cells equal to known account codes.
    For each code occurrence, capture the previous cell value (or 'null' if it's the first col).
    Returns list of (Account, Value) with Value numeric-coerced when possible.
    """
    soup = BeautifulSoup(html, 'html.parser')  # built-in parser; no lxml required
    tables = soup.find_all('table')
    if not tables:
        return []

    codes_set = set(account_codes)  # O(1) membership checks
    results: List[Tuple[str, object]] = []

    for table in tables:
        for row in table.find_all('tr'):
            cols = row.find_all(['td', 'th'])
            if not cols:
                continue
            row_text = [c.get_text(strip=True) for c in cols]
            if not row_text:
                continue

            # Any account codes present in this row?
            possible = set(row_text).intersection(codes_set)
            if not possible:
                continue

            # For each matched code, collect value from previous column (if exists)
            for code in possible:
                indices = [i for i, x in enumerate(row_text) if x == code]
                for idx in indices:
                    if idx > 0:
                        prev_val = parse_numeric(row_text[idx - 1])
                        results.append((code, prev_val))
                    else:
                        results.append((code, "null"))
    return results


# ==============================
# BUILDERS: Excel & ZIP outputs
# ==============================
def build_excel_bytes(per_pdf_sheets: List[Tuple[str, pd.DataFrame]],
                      log_df: pd.DataFrame,
                      include_consolidated: bool = True) -> bytes:
    """
    Combined workbook:
      - One sheet per PDF (sheet name derived from file name)
      - Optional 'Consolidated' sheet (all rows with Source column)
      - 'Log' sheet
    """
    bio = io.BytesIO()
    used_sheet_names = set()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        # One sheet per PDF
        for base_name, df in per_pdf_sheets:
            # Derive sheet name similar to your original (token after second underscore if present)
            parts = base_name.split("_")
            if len(parts) > 2:
                raw_name = os.path.splitext(parts[2])[0]
            else:
                raw_name = os.path.splitext(base_name)[0]
            sheet_name = make_unique_sheet_name(raw_name, used_sheet_names)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Optional consolidated
        if include_consolidated and per_pdf_sheets:
            rows = []
            for base_name, df in per_pdf_sheets:
                if not df.empty:
                    tmp = df.copy()
                    tmp.insert(0, "Source", base_name)
                    rows.append(tmp)
            if rows:
                consolidated = pd.concat(rows, ignore_index=True)
            else:
                consolidated = pd.DataFrame(columns=["Source", "Account", "Value"])
            cons_name = make_unique_sheet_name("Consolidated", used_sheet_names)
            consolidated.to_excel(writer, sheet_name=cons_name, index=False)

        # Log sheet
        log_name = make_unique_sheet_name("Log", used_sheet_names)
        log_df.to_excel(writer, sheet_name=log_name, index=False)
    return bio.getvalue()


def build_excel_bytes_for_single_sheet(df: pd.DataFrame,
                                       base_name_for_sheet: str,
                                       log_df: Optional[pd.DataFrame] = None) -> bytes:
    """
    Single-workbook for one PDF:
      - Sheet 1: the extracted rows (Account, Value), named after the PDF (sanitized)
      - Optional Sheet 2: Log (one-row log for that file)
    Caller controls the download filename to match the PDF base name.
    """
    bio = io.BytesIO()
    used_sheet_names = set()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        sheet_name = make_unique_sheet_name(base_name_for_sheet, used_sheet_names)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        if log_df is not None and not log_df.empty:
            log_sheet = make_unique_sheet_name("Log", used_sheet_names)
            log_df.to_excel(writer, sheet_name=log_sheet, index=False)
    return bio.getvalue()


def zip_excels_per_pdf(per_pdf_sheets: List[Tuple[str, pd.DataFrame]],
                       logs: List[Dict]) -> bytes:
    """
    Create a ZIP where each entry is an Excel workbook for a single PDF.
    The workbook contains one sheet named from the PDF (sanitized), and a Log sheet.
    The file name inside the ZIP is <PDF_BASENAME>.xlsx
    """
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Build a mapping from file name to its log dict for quick lookup
        log_map = {log.get("file"): log for log in logs}
        for (pdf_name, df) in per_pdf_sheets:
            base = os.path.splitext(os.path.basename(pdf_name))[0] or "file"
            log_df = pd.DataFrame([log_map.get(pdf_name, {})])
            xbio = io.BytesIO()
            used_sheet_names = set()
            with pd.ExcelWriter(xbio, engine="openpyxl") as writer:
                sheet_name = make_unique_sheet_name(base, used_sheet_names)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                log_sheet = make_unique_sheet_name("Log", used_sheet_names)
                log_df.to_excel(writer, sheet_name=log_sheet, index=False)
            zf.writestr(f"{base}.xlsx", xbio.getvalue())
    return bio.getvalue()


def to_zip_of_htmls(per_file_html: List[Tuple[str, str]]) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, html in per_file_html:
            safe = os.path.splitext(os.path.basename(fname))[0][:150] or "file"
            zf.writestr(f"{safe}.html", html.encode("utf-8", errors="ignore"))
    return bio.getvalue()


# ==============================
# SIDEBAR CONTROLS
# ==============================
with st.sidebar:
    st.header("Controls")
    input_mode = st.radio("Input source", ["Upload PDFs", "Pick from local folder"])
    include_consolidated = st.checkbox("Include Consolidated sheet (for combined workbook)", value=True)
    export_zip_per_pdf = st.checkbox(
        "Export one Excel per PDF (ZIP)",
        value=False,
        help="When multiple PDFs are processed, export a ZIP with <PDF>.xlsx for each."
    )
    include_html_zip = st.checkbox("Also provide ZIP of generated HTML files", value=False)
    preview_rows = st.number_input("Preview rows per sheet", min_value=5, max_value=200, value=20, step=5)
    st.caption("Note: Extractor matches cells equal to a known account code, capturing the **previous column** value.")


# ==============================
# INPUT (Upload or Local folder)
# ==============================
pdf_files: List[Tuple[str, bytes]] = []

if input_mode == "Upload PDFs":
    uploads = st.file_uploader("Upload one or more PDF files", type=["pdf"], accept_multiple_files=True)
    if uploads:
        for up in uploads:
            pdf_files.append((up.name, up.read()))
else:
    folder = st.text_input("Local folder path", value="", placeholder=r"C:\path\to\pdfs or ./pdfs")
    if folder:
        try:
            entries = []
            for fname in os.listdir(folder):
                path = os.path.join(folder, fname)
                if os.path.isfile(path) and fname.lower().endswith(".pdf"):
                    entries.append(path)
            entries.sort()
            if not entries:
                st.warning("No PDFs found in that folder.")
            else:
                picks = st.multiselect("Choose PDFs to process", options=entries, default=entries)
                for path in picks:
                    with open(path, "rb") as f:
                        pdf_files.append((os.path.basename(path), f.read()))
        except Exception as e:
            st.error(f"Could not list/read folder: {e}")

if not pdf_files:
    st.info("Add PDFs to begin.")
    st.stop()

st.success(f"Selected **{len(pdf_files)}** PDF(s).")
run = st.button("▶️ Convert & Extract", type="primary")


# ==============================
# RUN
# ==============================
if run:
    per_pdf_sheets: List[Tuple[str, pd.DataFrame]] = []
    logs: List[Dict] = []
    html_cache: List[Tuple[str, str]] = []
    progress = st.progress(0.0, text="Starting...")

    for idx, (name, data) in enumerate(pdf_files, start=1):
        progress.progress((idx - 1) / len(pdf_files), text=f"Processing {name} ({idx}/{len(pdf_files)})")
        log = {
            "file": name,
            "status": "ok",
            "exception": None,
            "html_chars": None,
            "rows_extracted": 0,
            "tables_found": None,
        }
        try:
            html_str = cached_pdf_to_html(name, data)
            html_cache.append((name, html_str))
            log["html_chars"] = len(html_str) if html_str else 0

            # Extract rows
            soup = BeautifulSoup(html_str, 'html.parser')
            tables = soup.find_all('table')
            log["tables_found"] = len(tables)

            rows = extract_rows_from_html(html_str)
            log["rows_extracted"] = len(rows)

            if rows:
                df = pd.DataFrame(rows, columns=['Account', 'Value'])
                per_pdf_sheets.append((name, df))
            else:
                # Match original behavior: skip empty sheets, but keep in log
                pass

        except Exception as e:
            log["status"] = "error"
            log["exception"] = "".join(traceback.format_exception(type(e), e, e.__traceback__))

        logs.append(log)

    progress.progress(1.0, text="Done")

    # Show log
    log_df = pd.DataFrame(logs)
    st.subheader("Log")
    st.dataframe(log_df, use_container_width=True, hide_index=True)

    # Previews for a quick sanity check
    if per_pdf_sheets:
        st.subheader("Previews (first rows)")
        for base_name, df in per_pdf_sheets:
            st.markdown(f"**{base_name}** — {len(df)} row(s)")
            st.dataframe(df.head(preview_rows), use_container_width=True, hide_index=True)
    else:
        st.warning("No data extracted from any PDF (no sheets to create).")

    st.divider()
    st.subheader("Download")

    # === Download logic based on your requested behavior ===
    if per_pdf_sheets:
        # A) Single PDF → name Excel exactly like the PDF
        if len(per_pdf_sheets) == 1 and not export_zip_per_pdf:
            single_name, single_df = per_pdf_sheets[0]
            pdf_base = os.path.splitext(os.path.basename(single_name))[0] or "file"
            # One-row log for this file
            single_log_df = pd.DataFrame([row for row in logs if row.get("file") == single_name])
            x_single = build_excel_bytes_for_single_sheet(single_df, pdf_base, log_df=single_log_df)
            st.download_button(
                label=f"⬇️ Download Excel ({pdf_base}.xlsx)",
                data=x_single,
                file_name=f"{pdf_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        # B) Multiple PDFs + ZIP per-PDF Excels
        elif export_zip_per_pdf and len(per_pdf_sheets) >= 1:
            z_bytes = zip_excels_per_pdf(per_pdf_sheets, logs)
            st.download_button(
                label="⬇️ Download ZIP (Excels per PDF)",
                data=z_bytes,
                file_name="Per_PDF_Excels.zip",
                mime="application/zip",
                type="primary",
            )

        # C) Default combined workbook
        else:
            x_bytes = build_excel_bytes(per_pdf_sheets, log_df, include_consolidated=include_consolidated)
            st.download_button(
                label="⬇️ Download Excel (Combined workbook)",
                data=x_bytes,
                file_name="Aggregated_5300_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
    else:
        st.info("No extracted rows to export.")

    # Optional: HTML ZIP for QA
    if include_html_zip and html_cache:
        z_html = to_zip_of_htmls(html_cache)
        st.download_button(
            label="⬇️ Download ZIP (Generated HTML files)",
            data=z_html,
            file_name="html_outputs.zip",
            mime="application/zip",
        )