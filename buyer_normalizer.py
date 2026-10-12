import re
from typing import Dict, Any, List, Tuple

GST_STATE_MAP = {
    '01': 'JAMMU AND KASHMIR', '02': 'HIMACHAL PRADESH', '03': 'PUNJAB', '04': 'CHANDIGARH',
    '05': 'UTTARAKHAND', '06': 'HARYANA', '07': 'DELHI', '08': 'RAJASTHAN', '09': 'UTTAR PRADESH',
    '10': 'BIHAR', '11': 'SIKKIM', '12': 'ARUNACHAL PRADESH', '13': 'NAGALAND', '14': 'MANIPUR',
    '15': 'MIZORAM', '16': 'TRIPURA', '17': 'MEGHALAYA', '18': 'ASSAM', '19': 'WEST BENGAL',
    '20': 'JHARKHAND', '21': 'ODISHA', '22': 'CHHATTISGARH', '23': 'MADHYA PRADESH', '24': 'GUJARAT',
    '26': 'DADRA AND NAGAR HAVELI AND DAMAN AND DIU', '27': 'MAHARASHTRA', '29': 'KARNATAKA',
    '30': 'GOA', '32': 'KERALA', '33': 'TAMIL NADU', '34': 'PUDUCHERRY', '36': 'TELANGANA', '37': 'ANDHRA PRADESH'
}

CITY_STATE_MAP = {
    'kolkata': ('WEST BENGAL', '19'), 'calcutta': ('WEST BENGAL', '19'), 'howrah': ('WEST BENGAL', '19'),
    'liluah': ('WEST BENGAL', '19'), 'uluberia': ('WEST BENGAL', '19'), 'murshidabad': ('WEST BENGAL', '19'),
    'salar': ('WEST BENGAL', '19'), 'singur': ('WEST BENGAL', '19'), 'jagatdal': ('WEST BENGAL', '19'),
    'dhuliyan': ('WEST BENGAL', '19'), 'birbhum': ('WEST BENGAL', '19'), 'murarai': ('WEST BENGAL', '19'),
    '24 parganas': ('WEST BENGAL', '19'), 'malipanchghara': ('WEST BENGAL', '19'), 'siliguri': ('WEST BENGAL', '19'),
    'ghusuri': ('WEST BENGAL', '19'),
    'west bengal': ('WEST BENGAL', '19'), 'w.b.': ('WEST BENGAL', '19'), 'w.b': ('WEST BENGAL', '19'),
    'guwahati': ('ASSAM', '18'), 'gauhati': ('ASSAM', '18'), 'bongaigaon': ('ASSAM', '18'),
    'paglasthan': ('ASSAM', '18'), 'barpeta': ('ASSAM', '18'), 'pathsala': ('ASSAM', '18'),
    'howly': ('ASSAM', '18'), 'lakhipur': ('ASSAM', '18'), 'fancy bazar': ('ASSAM', '18'), 'machkhowa': ('ASSAM', '18'),
    'assam': ('ASSAM', '18'),
    'muzaffarpur': ('BIHAR', '10'), 'bela': ('BIHAR', '10'), 'patna': ('BIHAR', '10'), 'bihar': ('BIHAR', '10'),
    'ghatsila': ('JHARKHAND', '20'), 'singhbhum': ('JHARKHAND', '20'), 'ranchi': ('JHARKHAND', '20'), 'jharkhand': ('JHARKHAND', '20'),
    'dharmanagar': ('TRIPURA', '16'), 'teliamura': ('TRIPURA', '16'), 'khowai': ('TRIPURA', '16'), 'tripura': ('TRIPURA', '16'),
    'odisha': ('ODISHA', '21'), 'cuttack': ('ODISHA', '21'), 'bhubaneswar': ('ODISHA', '21'),
    'raipur': ('CHHATTISGARH', '22'), 'chhattisgarh': ('CHHATTISGARH', '22')
}


def clean_line_phone_email(line: str) -> str:
    """
    Strip email addresses, phone/mobile/fax numbers from an address line.
    Preserves 6-digit postal PIN codes (e.g. 711315, 781001, PIN - 743125).
    """
    s = (line or '').strip()
    if not s:
        return ''

    # 1. Strip labeled email or raw email (including optional contact:, mailto:, etc.)
    s = re.sub(r'(?i)\b(email|e-mail|mail|contact)\s*[:.\-]?\s*[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', '', s)
    s = re.sub(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', '', s)

    # 2. Strip labeled phone/fax numbers (Ph:, Mob:, Mobile:, Tel:, Phone:, Contact:, Call:, M:, Fax:, etc.)
    s = re.sub(r'(?i)\b(ph(\.|\s*no)?|mob(ile)?|phone|tel(ephone)?|contact|cell|call|fax|m)\s*[:.\-/#]?\s*(\+?91[\-\s]?)?[\d\s\-/()]{7,}', '', s)

    # 3. Strip standalone 10-digit mobile numbers with optional +91 and optional hyphen/space separator
    # e.g. 9830012345, 98300-12345, 98300 12345, +91-98300-12345
    s = re.sub(r'(\+?91[\-\s]?)?[6-9]\d{4}[\-\s]?\d{5}\b', '', s)
    s = re.sub(r'(\+?91[\-\s]?)?[6-9]\d{9}\b', '', s)

    # 4. Strip standard landline/fax numbers with STD code (033-26541234 or 0361-2541234)
    s = re.sub(r'\b0\d{2,4}[\-\s]?\d{6,8}\b', '', s)

    # 5. If only leftover labels like 'Email:', 'Ph:', 'Fax:', 'Contact:' remain
    if re.match(r'(?i)^\s*(email|e-mail|mail|ph(\.|\s*no)?|phone|mob(ile)?|tel(ephone)?|contact|call|cell|fax|m)\s*[:.\-/#]?\s*$', s):
        return ''

    # 6. Clean up multiple slashes, commas, trailing punctuation
    s = re.sub(r'\s*[/,]\s*[/,]+', ',', s)
    s = re.sub(r'[\s,:\-/]+$', '', s)
    s = re.sub(r'^[\s,:\-/]+', '', s)
    return s.strip()


def normalize_buyer_profile(p: Dict[str, Any]) -> Dict[str, Any]:
    """
    Standardize a buyer profile for Shakambhari Enterprises:
    1. Canonical lines:
       - Line 1: 'Buyer :'
       - Line 2: <Buyer Name>
       - Lines 3..k: Address lines (stripped of phone numbers, emails, and embedded state codes)
       - Next to last: 'STATE- <STATE_NAME> CODE-<STATE_CODE>' (mandatory unless otherwise mentioned)
       - Last: 'GSTIN - <GSTIN>' or 'PAN NO. - ...'
    2. Infers state code and state name from GSTIN first 2 digits or city/region fallback.
    3. Normalizes top-level 'gstin', 'state', 'state_code', and 'default_tax_type' (CGST_SGST for 19, IGST for others).
    """
    buyer_name = (p.get('buyer_name') or '').strip()
    raw_lines = p.get('buyer_details') or []
    if isinstance(raw_lines, str):
        raw_lines = [line.strip() for line in raw_lines.splitlines() if line.strip()]

    gstin = (p.get('gstin') or '').strip().upper()

    # Step 1: Detect and extract GSTIN if present in lines
    for l in raw_lines:
        clean_l = re.sub(r'[\s\xa0\u00a0\ufffd]+', ' ', str(l))
        m = re.search(r'\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b', clean_l.upper())
        if m:
            gstin = m.group(1)
            break

    # Step 2: Detect state and state code
    state_code = (p.get('state_code') or '').strip()
    state_name = ''

    # Check GSTIN first
    if gstin and len(gstin) >= 2 and gstin[:2].isdigit():
        code = gstin[:2]
        if code in GST_STATE_MAP:
            state_code = code
            state_name = GST_STATE_MAP[code]

    # Check lines for existing state code
    if not state_code:
        for l in raw_lines:
            clean_l = re.sub(r'[\s\xa0\u00a0\ufffd]+', ' ', str(l))
            m = re.search(r'(?i)\bstate\s*[-:]?\s*([A-Za-z\s]+?)\s*[-:]?\s*code\s*[-:]?\s*(\d{1,2})\b', clean_l)
            if m:
                c = m.group(2).zfill(2)
                if c in GST_STATE_MAP:
                    state_code = c
                    state_name = GST_STATE_MAP[c]
                    break
            m2 = re.search(r'(?i)\b([A-Za-z\s]+?)\s*code\s*[:\-]?\s*(\d{1,2})\b', clean_l)
            if m2:
                c = m2.group(2).zfill(2)
                if c in GST_STATE_MAP:
                    state_code = c
                    state_name = GST_STATE_MAP[c]
                    break
            m3 = re.search(r'(?i)\bstate\s*code\s*[:\-]?\s*(\d{1,2})\b', clean_l)
            if m3:
                c = m3.group(1).zfill(2)
                if c in GST_STATE_MAP:
                    state_code = c
                    state_name = GST_STATE_MAP[c]
                    break

    # Fallback to city/state keywords
    if not state_code:
        full_text = ' '.join(str(x) for x in raw_lines).lower() + ' ' + (p.get('state') or '').lower() + ' ' + buyer_name.lower()
        for kw, (s_name, s_code) in CITY_STATE_MAP.items():
            if re.search(r'\b' + re.escape(kw) + r'\b', full_text):
                state_code = s_code
                state_name = s_name
                break

    if state_code and not state_name and state_code in GST_STATE_MAP:
        state_name = GST_STATE_MAP[state_code]

    # Step 3: Clean address lines (strip phone, email, and embedded state codes)
    cleaned_address_lines = []
    pan_line = ''
    for l in raw_lines:
        clean_l = re.sub(r'[\s\xa0\u00a0\ufffd]+', ' ', str(l)).strip()
        lowered = clean_l.lower()
        # Skip Buyer: header
        if lowered in {'buyer:', 'buyer :'}:
            continue
        # Skip buyer name if duplicate
        if lowered == buyer_name.lower():
            continue
        # Check if PAN line
        if re.search(r'(?i)\bpan\s*no\.?\s*[-:]\s*([A-Z]{5}[0-9]{4}[A-Z]{1})\b', clean_l):
            pan_line = clean_l
            continue
        # Check if GST line
        if re.search(r'(?i)\bgst(in|\s*no)?\b', clean_l) or (gstin and gstin in clean_l.upper()):
            continue
        # Check if line was a dedicated state code line (e.g. 'STATE- ASSAM CODE-18', 'Tripura Code : 16', 'STATE CODE - 20')
        if re.fullmatch(r'(?i)^\s*(state\s*[-:]?)?[A-Za-z\s]*code\s*[:\-]?\s*\d{1,2}\s*$', clean_l) or re.fullmatch(r'(?i)^\s*state\s*[-:]?\s*[A-Za-z\s]*\s*$', clean_l):
            continue
        if re.search(r'(?i)^\s*state\s*[-:]?\s*[A-Za-z\s]+[-:]?\s*code\s*[-:]?\s*\d{1,2}\s*$', clean_l):
            continue

        # Strip embedded CODE - XX at the end of town lines (e.g. DHARMANAGAR, TRIPURA-799250 CODE-16)
        clean_l = re.sub(r'(?i)[,\s\-]+code\s*[-:]?\s*\d{1,2}\b', '', clean_l).strip()

        # If line is just 'STATE' or matches a known state name, skip it since canonical STATE line is added
        if clean_l.upper() in {'STATE', 'STATE:', 'STATE -'} or clean_l.upper() in GST_STATE_MAP.values():
            continue

        # Strip phone & email
        clean_l = clean_line_phone_email(clean_l)
        if clean_l and clean_l.lower() != buyer_name.lower():
            cleaned_address_lines.append(clean_l)

    # Step 4: Assemble canonical lines
    final_lines = ['Buyer :']
    if buyer_name and buyer_name != 'Buyer :':
        final_lines.append(buyer_name)
    final_lines.extend(cleaned_address_lines)

    # Add canonical state code line if state_code exists
    if state_code and state_name:
        final_lines.append(f'STATE- {state_name} CODE-{state_code}')

    # Add GSTIN or PAN line
    if gstin:
        final_lines.append(f'GSTIN - {gstin}')
    elif pan_line:
        final_lines.append(pan_line)

    tax_type = 'CGST_SGST' if state_code == '19' else 'IGST'

    return {
        'profile_id': p.get('profile_id'),
        'buyer_name': buyer_name,
        'buyer_details': final_lines,
        'gstin': gstin,
        'state': state_name.title() if state_name else p.get('state', ''),
        'state_code': state_code or p.get('state_code', ''),
        'default_tax_type': tax_type
    }
