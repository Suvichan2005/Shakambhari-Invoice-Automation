"""
Shakambhari Enterprises Invoice Generator - Cloud Version
==========================================================
A Flask-based invoice generation system deployed on Google Cloud with:
- Google Sheets for data storage (buyers, transport modes, invoice records)
- Google Cloud Storage for file storage (Excel, PDF invoices)
- WeasyPrint for PDF generation (no Windows dependency)
"""

import os
import io
import re
import json
import base64
import tempfile
import time
import shutil
import subprocess
import hashlib
import hmac
from copy import copy
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict, deque
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, Response, session, send_from_directory, make_response
from datetime import datetime, timezone, timedelta
try:
    from zoneinfo import ZoneInfo
    IST = ZoneInfo('Asia/Kolkata')
    datetime.now(IST)
except Exception:
    IST = timezone(timedelta(hours=5, minutes=30), name='IST')

def get_ist_now():
    """Return current datetime in Indian Standard Time (Asia/Kolkata)."""
    return datetime.now(IST)

def get_ist_today_str():
    """Return today's date formatted as YYYY-MM-DD in IST."""
    return get_ist_now().strftime('%Y-%m-%d')
import uuid
from num2words import num2words
from typing import Any, List, Dict, Optional, Tuple
from werkzeug.exceptions import HTTPException
import urllib.request
import urllib.error
from pathlib import Path
import sys

_cloud_dir = os.path.dirname(os.path.abspath(__file__))
if _cloud_dir not in sys.path:
    sys.path.insert(0, _cloud_dir)

# Cloud integrations
from sheets_db import GoogleSheetsDB, init_sheets_db
from cloud_storage import CloudStorage, init_cloud_storage
from settings_manager import get_settings, update_settings, DEFAULT_SETTINGS


def sanitize_settings_for_client(settings_dict: Dict[str, Any]) -> Dict[str, Any]:
    """Strip server-only credentials (e.g. Gemini backend API key) before responding to browser."""
    safe = dict(settings_dict or {})
    backend_key = safe.pop('gemini_backend_api_key', '')
    safe['has_backend_api_key'] = bool(backend_key)
    return safe
from buyer_normalizer import (
    normalize_buyer_profile,
    clean_line_phone_email,
    GST_STATE_MAP,
    CITY_STATE_MAP
)

# Excel handling
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# PDF generation (cloud-compatible)
try:
    from weasyprint import HTML, CSS
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False
    print("WARNING: WeasyPrint not available. PDF generation will be skipped.")

app = Flask(__name__)

# Secret key for Flask sessions; prefer explicit env var in production.
_secret_key = os.environ.get('FLASK_SECRET_KEY') or os.environ.get('SECRET_KEY')
if not _secret_key:
    _secret_key = os.urandom(32).hex()
    print("WARNING: FLASK_SECRET_KEY is not set. Using ephemeral key; sessions reset on restart.")

app.secret_key = _secret_key
app.permanent_session_lifetime = timedelta(days=30)
# Firebase Hosting CDN only passes through cookies named '__session'.
app.config['SESSION_COOKIE_NAME'] = '__session'
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

secure_cookie_env = os.environ.get('SESSION_COOKIE_SECURE', '').strip().lower()
if secure_cookie_env in {'1', 'true', 'yes'}:
    app.config['SESSION_COOKIE_SECURE'] = True
elif secure_cookie_env in {'0', 'false', 'no'}:
    app.config['SESSION_COOKIE_SECURE'] = False
else:
    # Default to secure only when running in managed cloud runtime.
    app.config['SESSION_COOKIE_SECURE'] = bool(os.environ.get('K_SERVICE'))

# Application authentication: PBKDF2 salted hash to avoid keeping raw password in memory
_AUTH_SALT = b"shakambhari_auth_v1_salt"

def hash_app_password(password: str) -> str:
    """Generate deterministic PBKDF2 hash of password with fixed salt."""
    return hashlib.pbkdf2_hmac(
        'sha256', password.strip().encode('utf-8'), _AUTH_SALT, 100000
    ).hex()

def set_app_password(password: str) -> None:
    """Set or update app password hash in-memory."""
    cleaned = password.strip() if password else ''
    if cleaned:
        app.config['APP_PASSWORD_HASH'] = hash_app_password(cleaned)
        app.config['AUTH_ENABLED'] = True
    else:
        app.config['APP_PASSWORD_HASH'] = None
        app.config['AUTH_ENABLED'] = False

def verify_app_password(entered_password: str) -> bool:
    """Constant-time verification of entered password against PBKDF2 hash."""
    if not app.config.get('AUTH_ENABLED'):
        return True
    stored_hash = app.config.get('APP_PASSWORD_HASH')
    if not stored_hash or not entered_password:
        return False
    entered_hash = hash_app_password(entered_password)
    return hmac.compare_digest(entered_hash, stored_hash)

_raw_app_password = os.environ.get('APP_PASSWORD', '').strip()
set_app_password(_raw_app_password)
del _raw_app_password

# Basic in-memory per-IP rate limiting for sensitive routes.
RATE_LIMITS = {
    'login': (5, 60),  # Stricter limit: 5 attempts per minute to mitigate brute-forcing
    'generate_invoice': (30, 60),
    'calculate_preview': (300, 60),
    'api_get_invoice': (180, 60),
    'download_xlsx': (120, 60),
    'download_pdf': (120, 60),
}
_request_windows: Dict[str, deque] = defaultdict(deque)

# Initialize cloud services (lazy loading)
_sheets_db: Optional[GoogleSheetsDB] = None
_cloud_storage: Optional[CloudStorage] = None


AUTH_EXEMPT_ENDPOINTS = {
    'login',
    'logout',
    'health_check',
    'favicon',
    'static',
}


@app.before_request
def require_authentication():
    """Protect routes with a simple session login when APP_PASSWORD is set."""
    if not app.config.get('AUTH_ENABLED'):
        return None

    endpoint = request.endpoint or ''
    if endpoint in AUTH_EXEMPT_ENDPOINTS or endpoint.startswith('static'):
        return None

    if session.get('authenticated'):
        return None

    if request.path.startswith('/api/'):
        return jsonify({'error': 'Authentication required'}), 401

    flash('Please login to continue.', 'warning')
    return redirect(url_for('login', next=request.full_path if request.query_string else request.path))


def _normalize_post_login_target(next_url: str) -> str:
    """Return a safe GET endpoint for post-login navigation."""
    target = (next_url or '').strip()
    if not target.startswith('/'):
        return url_for('index')

    # Never redirect to API or known POST-only endpoints after login.
    if target.startswith('/api/') or target.startswith('/generate_invoice') or target.startswith('/logout'):
        return url_for('index')

    return target


@app.before_request
def enforce_rate_limit():
    """Apply lightweight per-IP throttling to reduce abuse spikes."""
    endpoint = request.endpoint or ''
    if endpoint not in RATE_LIMITS:
        return None

    limit, window_seconds = RATE_LIMITS[endpoint]
    remote = request.headers.get('X-Forwarded-For', request.remote_addr or 'unknown').split(',')[0].strip()
    now = time.time()
    key = f"{remote}:{endpoint}"
    bucket = _request_windows[key]

    while bucket and bucket[0] <= now - window_seconds:
        bucket.popleft()

    if len(bucket) >= limit:
        if request.path.startswith('/api/'):
            return jsonify({'error': 'Too many requests. Please retry shortly.'}), 429
        return render_template('error.html',
                              code=429,
                              title='Too Many Requests',
                              message='Please wait a few seconds and try again.'), 429

    bucket.append(now)
    return None


@app.after_request
def add_security_headers(response: Response) -> Response:
    """Inject standard defensive HTTP security headers into every response."""
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'

    # Prevent caching of sensitive financial/API views and authenticated sessions
    if not request.path.startswith('/static/'):
        if request.path.startswith('/api/') or session.get('authenticated'):
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'

    return response


@app.route('/login', methods=['GET', 'POST'])
def login():
    """Login route for lightweight app-level protection."""
    if not app.config.get('AUTH_ENABLED'):
        session.clear()
        session.permanent = True
        session['authenticated'] = True
        return redirect(url_for('index'))

    if request.method == 'POST':
        password = request.form.get('password', '')
        if verify_app_password(password):
            session.clear()
            session.permanent = True
            session['authenticated'] = True
            flash('Login successful.', 'success')
            next_url = _normalize_post_login_target(
                request.args.get('next') or request.form.get('next') or url_for('index')
            )
            return redirect(next_url)

        # Anti-brute-force penalty delay to slow down dictionary attacks
        time.sleep(1.0)
        flash('Invalid password.', 'error')

    return render_template('login.html',
                          auth_enabled=app.config.get('AUTH_ENABLED', False),
                          next_url=request.args.get('next', url_for('index')))


@app.route('/logout', methods=['POST'])
def logout():
    """Clear login session."""
    session.clear()
    flash('Logged out successfully.', 'info')
    return redirect(url_for('login'))


def get_sheets_db() -> GoogleSheetsDB:
    """Get or initialize the Google Sheets database connection."""
    global _sheets_db
    if _sheets_db is None:
        _sheets_db = init_sheets_db()
    return _sheets_db


def get_sheets_db_or_none() -> Optional[GoogleSheetsDB]:
    """Return the Sheets DB if configured, otherwise None."""
    try:
        return get_sheets_db()
    except Exception as exc:
        app.logger.warning("Sheets DB unavailable: %s", exc)
        return None


def get_cloud_storage() -> Optional[CloudStorage]:
    """Get or initialize the Cloud Storage connection, returning None if unavailable."""
    global _cloud_storage
    if _cloud_storage is None:
        try:
            _cloud_storage = init_cloud_storage()
        except Exception as exc:
            app.logger.warning("Cloud Storage unavailable: %s", exc)
            return None
    return _cloud_storage


def _safe_filename(filename: str, expected_ext: str) -> str:
    """Normalize and validate download filenames to avoid path manipulation."""
    safe = os.path.basename((filename or '').strip())
    safe = re.sub(r'[^A-Za-z0-9._-]', '_', safe)
    ext = expected_ext.lower()
    if not safe.lower().endswith(ext):
        safe = f"{os.path.splitext(safe)[0]}{ext}"
    return safe


def _split_detail_lines(raw_text: str) -> List[str]:
    """Split a textarea value into cleaned non-empty lines."""
    return [line.strip() for line in (raw_text or '').splitlines() if line.strip()]


def _extract_gstin_from_lines(lines: List[str]) -> str:
    """Extract a GSTIN from detail lines when one is present."""
    for line in lines:
        match = re.search(r'(?:GST\s*IN|GSTIN)\s*[-:]\s*([0-9A-Z]{15})', line, re.IGNORECASE)
        if match:
            return match.group(1).upper()

        compact = re.sub(r'[^0-9A-Z]', '', line.upper())
        if len(compact) == 15 and re.fullmatch(r'[0-9A-Z]{15}', compact):
            return compact

    return ''


def _normalize_buyer_details_for_storage(buyer_name: str, raw_text: str, fallback_gstin: str = '') -> tuple[list[str], str]:
    """Build the canonical buyer_details block stored in Sheets."""
    lines = _split_detail_lines(raw_text)
    profile_data = {
        'buyer_name': buyer_name,
        'buyer_details': lines,
        'gstin': fallback_gstin
    }
    norm = normalize_buyer_profile(profile_data)
    return norm['buyer_details'], norm['gstin']


def _details_for_edit_form(profile: Dict[str, Any]) -> str:
    """Convert stored canonical buyer_details back into editable address-only lines."""
    buyer_name = (profile.get('buyer_name') or '').strip().lower()
    lines = profile.get('buyer_details') or []
    fallback_gstin = (profile.get('gstin') or '').strip().upper()
    cleaned = []

    for line in lines:
        stripped = (line or '').strip()
        lowered = stripped.lower()
        if not stripped:
            continue
        if lowered in {'buyer:', 'buyer :'}:
            continue
        if buyer_name and lowered == buyer_name:
            continue
        cleaned.append(stripped)

    if fallback_gstin and not any(fallback_gstin in line.upper() for line in cleaned):
        cleaned.append(f'GSTIN - {fallback_gstin}')

    return '\n'.join(cleaned)


def _filename_from_storage_url(storage_url: str, expected_ext: str) -> str:
    """Extract object name from a gs:// URL and normalize extension."""
    if not storage_url:
        return ''
    fname = storage_url.rsplit('/', 1)[-1]
    return _safe_filename(fname, expected_ext)


def _format_datetime_display(raw_value: str) -> str:
    """Format common timestamp/date strings for UI display."""
    if not raw_value:
        return ''
    text = str(raw_value).strip()
    for fmt in ('%Y-%m-%dT%H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d'):
        try:
            dt = datetime.strptime(text, fmt)
            if fmt == '%Y-%m-%d':
                return dt.strftime('%Y-%m-%d')
            return dt.strftime('%Y-%m-%d %H:%M')
        except ValueError:
            continue
    return text


def _invoice_sort_key(rec: Dict) -> Tuple[str, int, str]:
    """Sort key prioritizing invoice date ISO, numeric invoice number, and timestamp."""
    d = str(rec.get('invoice_date') or '').strip()
    m = re.match(r'^(\d{1,2})[/\.-](\d{1,2})[/\.-](\d{4})', d)
    if m:
        d = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    num = 0
    nm = re.search(r'(\d+)', str(rec.get('invoice_number') or ''))
    if nm:
        num = int(nm.group(1))
    created = str(rec.get('created_at') or '')
    return (d, num, created)


def _build_invoice_rows(records: List[Dict]) -> List[Dict]:
    """Normalize invoice records for index/dashboard displays."""
    rows: List[Dict] = []
    for rec in records:
        file_url = rec.get('file_url', '')
        pdf_url = rec.get('pdf_url', '')
        xlsx_name = _filename_from_storage_url(file_url, '.xlsx') if file_url else ''
        pdf_name = _filename_from_storage_url(pdf_url, '.pdf') if pdf_url else ''
        items = rec.get('items', [])
        if not isinstance(items, list):
            items = []

        created_raw = rec.get('created_at') or rec.get('invoice_date', '')
        modified_date = _format_datetime_display(created_raw)
        sort_ts = 0
        for fmt in ('%Y-%m-%dT%H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d'):
            try:
                sort_ts = int(datetime.strptime(str(created_raw), fmt).timestamp())
                break
            except Exception:
                continue

        rows.append({
            **rec,
            'filename': xlsx_name,
            'pdf_filename': pdf_name,
            'items_count': len(items),
            'modified_date': modified_date,
            'sort_ts': sort_ts,
            'transport_mode': extract_transport_core(rec.get('transport_mode', '')),
        })

    rows.sort(key=_invoice_sort_key, reverse=True)
    return rows


def _parse_invoice_filename(filename: str) -> Dict[str, str]:
    """Infer invoice number and buyer label from generated invoice filenames."""
    stem = os.path.splitext(filename or '')[0]
    stem = re.sub(r'__\d+$', '', stem)

    if stem.lower().startswith('invoice_'):
        rest = stem[len('invoice_'):]
    else:
        rest = stem

    parts = [p for p in rest.split('_') if p]
    if not parts:
        return {'invoice_number': '', 'buyer_name': ''}

    invoice_number = ''
    buyer_tokens: List[str] = []

    # Invoice_023_2025_26_Prabhat_Aluminium_Industries
    if len(parts) >= 3 and parts[0].isdigit() and re.fullmatch(r'\d{4}', parts[1]) and re.fullmatch(r'\d{2}', parts[2]):
        invoice_number = f"{int(parts[0])}/{parts[1]}-{parts[2]}"
        buyer_tokens = parts[3:]
    # Invoice_8-2026-27_2026-04-17_Buyer_Name or Invoice_1-2025-26_...
    elif re.fullmatch(r'\d+-\d{4}-\d{2}', parts[0] or ''):
        m = re.match(r'^(\d+)-(\d{4})-(\d{2})$', parts[0])
        if m:
            invoice_number = f"{int(m.group(1))}/{m.group(2)}-{m.group(3)}"
        buyer_tokens = parts[1:]
    # Invoice_001_2026_27_Tirupati_Udyog (already handled above) or fallback first token only
    elif parts[0].isdigit():
        invoice_number = str(int(parts[0]))
        buyer_tokens = parts[1:]
    else:
        buyer_tokens = parts

    # Skip date token if present after invoice token.
    if buyer_tokens and re.fullmatch(r'\d{4}-\d{2}-\d{2}', buyer_tokens[0]):
        buyer_tokens = buyer_tokens[1:]

    buyer_name = ' '.join(buyer_tokens).replace('-', ' ').strip()
    buyer_name = re.sub(r'\s+', ' ', buyer_name)

    return {
        'invoice_number': invoice_number,
        'buyer_name': buyer_name,
    }


def _extract_invoice_data_from_xlsx_bytes(file_bytes: bytes, filename: str = '') -> Dict[str, Any]:
    """Extract invoice payload from an XLSX file stored in GCS or locally."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    try:
        sheet = wb.active

        invoice_number = ''
        invoice_date = ''
        ewaybill_number = ''
        ewaybill_date = ''
        transport_mode = ''
        delivery_charge = 0.0
        tax_type = 'IGST'
        subtotal = 0.0
        tax_amount = 0.0
        total_amount = 0.0
        digitally_signed = False

        if hasattr(sheet, '_images') and sheet._images:
            digitally_signed = True

        # Helper to parse dates flexibly
        def _parse_date_value(val_any: Any) -> str:
            if not val_any:
                return ''
            if hasattr(val_any, 'strftime'):
                return val_any.strftime('%Y-%m-%d')
            s = str(val_any).strip()
            m = re.search(r'(?i)\bdate\s*[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', s)
            if not m:
                m = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b', s)
            if m:
                raw_d = m.group(1).replace('-', '/')
                for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y/%m/%d', '%m/%d/%Y'):
                    try:
                        return datetime.strptime(raw_d, fmt).strftime('%Y-%m-%d')
                    except ValueError:
                        pass
            return ''

        # 1. Direct cell extraction for standard 2026-27 & older layouts
        f2_val = sheet['F2'].value
        if f2_val:
            invoice_date = _parse_date_value(f2_val)
        if not invoice_date:
            h2_val = sheet['H2'].value
            if h2_val:
                invoice_date = _parse_date_value(h2_val)

        a2_val = str(sheet['A2'].value or '').strip()
        m_inv = re.search(r'(?i)invoice\s*no\.?\s*[:\s]*([0-9a-zA-Z\-_/]+)', a2_val)
        if m_inv:
            invoice_number = m_inv.group(1).strip()
        else:
            e2_val = str(sheet['E2'].value or '').strip()
            m_inv = re.search(r'(?i)invoice\s*no\.?\s*[:\s]*([0-9a-zA-Z\-_/]+)', e2_val)
            if m_inv:
                invoice_number = m_inv.group(1).strip()

        a3_val = str(sheet['A3'].value or '').strip()
        m_ew = re.search(r'(?i)ewaybill\s*no\.?\s*[:\s]*(\d+)', a3_val)
        if m_ew:
            ewaybill_number = m_ew.group(1).strip()

        f3_val = sheet['F3'].value
        if f3_val:
            m_ewd = re.search(r'(?i)ewaybill\s*date\s*[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', str(f3_val))
            if m_ewd:
                ewaybill_date = _parse_date_value(m_ewd.group(1))

        a20_val = str(sheet['A20'].value or '').strip()
        m_trans = re.search(r'(?i)mode\s*of\s*transports?\s*[:\s]*(.*)', a20_val)
        if m_trans and m_trans.group(1).strip():
            transport_mode = extract_transport_core(m_trans.group(1).strip())

        # 2. General scan fallback
        for row in sheet.iter_rows(values_only=False):
            for cell in row:
                val = str(cell.value or '').strip()
                if not val:
                    continue

                if not invoice_number:
                    m = re.search(r'(?i)invoice\s*no\.?\s*[:\s]*([0-9a-zA-Z\-_/]+)', val)
                    if m:
                        invoice_number = m.group(1).strip()

                if not invoice_date and 'ewaybill' not in val.lower():
                    m = re.search(r'(?i)\bdate\s*[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', val)
                    if m:
                        invoice_date = _parse_date_value(m.group(1))

                if not ewaybill_number:
                    m = re.search(r'(?i)ewaybill\s*no\.?\s*[:\s]*(\d+)', val)
                    if m:
                        ewaybill_number = m.group(1).strip()

                if not ewaybill_date and 'ewaybill' in val.lower():
                    m = re.search(r'(?i)ewaybill\s*date\s*[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', val)
                    if m:
                        ewaybill_date = _parse_date_value(m.group(1))

                if not transport_mode:
                    m = re.search(r'(?i)mode\s*of\s*transports?\s*[:\s]*(.*)', val)
                    if m:
                        raw_t = m.group(1).strip()
                        if not raw_t:
                            next_val = sheet.cell(row=cell.row, column=cell.column+1).value or sheet.cell(row=cell.row, column=cell.column+2).value
                            if next_val:
                                raw_t = str(next_val).strip()
                        transport_mode = extract_transport_core(raw_t)

                if 'delivery charge' in val.lower() or val.lower() == 'delivery':
                    val_cell = sheet.cell(row=cell.row, column=9).value
                    delivery_charge = safe_float(val_cell, 0.0)

                if 's.g.s.t' in val.lower() or ('c.g.s.t' in val.lower() and '0.00%' not in str(sheet.cell(row=cell.row, column=5).value or '')):
                    rate_val = safe_float(sheet.cell(row=cell.row, column=5).value, 0.0)
                    amt_val = safe_float(sheet.cell(row=cell.row, column=9).value, 0.0)
                    if rate_val > 0 or amt_val > 0:
                        tax_type = 'CGST_SGST'

                if val.upper() == 'TOTAL':
                    amt = safe_float(sheet.cell(row=cell.row, column=9).value, 0.0)
                    if amt > 0:
                        total_amount = amt

                if 'digitally signed' in val.lower():
                    digitally_signed = True

        buyer_details = []
        buyer_start_row = 13
        for r in range(1, 20):
            cval = str(sheet[f'A{r}'].value or '').strip()
            if 'BUYER' in cval.upper():
                buyer_start_row = r + 1
                break

        for r in range(buyer_start_row, buyer_start_row + 7):
            v = sheet[f'A{r}'].value
            if v and not str(v).upper().startswith('MODE OF TRANSPORT'):
                buyer_details.append(str(v).strip())

        buyer_name = buyer_details[0] if buyer_details else ''
        buyer_gstin = ''
        for line in buyer_details:
            m = re.search(r'GSTIN\s*[-:]\s*([A-Z0-9]+)', str(line), re.IGNORECASE)
            if m:
                buyer_gstin = m.group(1).upper()
                break

        ship_from_details = []
        for r in range(5, 10):
            v = sheet[f'F{r}'].value
            if v and not str(v).upper().startswith('SHIP TO'):
                ship_from_details.append(str(v).strip())

        ship_to_details = []
        for r in range(13, 19):
            v = sheet[f'F{r}'].value
            if v:
                ship_to_details.append(str(v).strip())

        # Extract items
        items = []
        item_start_row = 22 if str(sheet['Z1'].value or '') == 'v2' or sheet['A21'].value else 18
        for row in range(item_start_row, item_start_row + 10):
            desc_raw = str(sheet[f'A{row}'].value or '').strip()
            hsn_val = sheet[f'H{row}'].value or sheet[f'B{row}'].value
            hsn = str(hsn_val).strip() if hsn_val not in (None, '') else ''
            qty = safe_float(sheet[f'F{row}'].value, 0.0)
            rate = safe_float(sheet[f'G{row}'].value, 0.0)

            if desc_raw or qty or rate:
                base_desc = re.sub(r'^\d+\.\s*', '', desc_raw).strip()
                bags = ''
                bag_match = re.search(r'\((\d+(?:\.\d+)?)\s*Bags?\)', base_desc, re.IGNORECASE)
                if bag_match:
                    bags = bag_match.group(1)
                    base_desc = re.sub(r'\s*\(\d+(?:\.\d+)?\s*Bags?\)', '', base_desc, flags=re.IGNORECASE).strip()

                items.append({
                    'description': base_desc,
                    'bags': bags,
                    'hsn': hsn,
                    'quantity': qty,
                    'rate': rate,
                })

        parsed_from_name = _parse_invoice_filename(filename)
        if not invoice_number:
            invoice_number = parsed_from_name.get('invoice_number', '')
        if not buyer_name:
            buyer_name = parsed_from_name.get('buyer_name', '')

        return {
            'invoice_number': invoice_number,
            'invoice_date': invoice_date,
            'buyer_name': buyer_name,
            'buyer_gstin': buyer_gstin,
            'buyer_details': buyer_details,
            'transport_mode': transport_mode,
            'delivery_charge': delivery_charge,
            'ship_from_details': ship_from_details,
            'ship_to_enabled': bool(ship_to_details),
            'ship_to_details': ship_to_details,
            'ewaybill_number': ewaybill_number,
            'ewaybill_date': ewaybill_date,
            'digitally_signed': digitally_signed,
            'items': items,
            'tax_type': tax_type,
            'filename': filename,
            'total_amount': total_amount,
            'pdf_filename': filename.replace('.xlsx', '.pdf') if filename.lower().endswith('.xlsx') else '',
        }
    finally:
        wb.close()


def _merge_with_storage_rows(sheet_rows: List[Dict], storage_rows: List[Dict]) -> List[Dict]:
    """Merge invoice metadata with bucket files so the modal shows all available files."""
    merged = list(sheet_rows)
    existing = {row.get('filename', '') for row in sheet_rows if row.get('filename')}

    for blob in storage_rows:
        filename = blob.get('filename', '')
        if not filename or filename in existing:
            continue

        parsed = _parse_invoice_filename(filename)
        updated = blob.get('updated')
        modified_date = ''
        if updated:
            try:
                modified_date = updated.strftime('%Y-%m-%d %H:%M')
            except Exception:
                modified_date = str(updated)

        merged.append({
            'filename': filename,
            'pdf_filename': filename.replace('.xlsx', '.pdf'),
            'invoice_number': parsed.get('invoice_number', ''),
            'buyer_name': parsed.get('buyer_name', ''),
            'modified_date': modified_date,
            'sort_ts': int(updated.timestamp()) if updated else 0,
            'items_count': 0,
            'tax_type': '',
            'transport_mode': '',
            'total_amount': '',
        })

    merged.sort(key=_invoice_sort_key, reverse=True)
    return merged


# ===================== UTILITY FUNCTIONS =====================

def _financial_year_suffix(today: datetime = None) -> str:
    """Get the financial year suffix like /2025-26."""
    today = today or get_ist_now()
    year = today.year
    if today.month >= 4:
        start = year
        end = year + 1
    else:
        start = year - 1
        end = year
    return f"/{start}-{str(end)[-2:]}"


def suggest_next_invoice_number() -> str:
    """Suggest the next sequential invoice number based on highest number in current FY."""
    fy = _financial_year_suffix()
    db = get_sheets_db_or_none()
    storage = get_cloud_storage()
    
    max_num = 0
    all_invoices = []
    
    if db is not None:
        try:
            all_invoices.extend(db.get_all_invoices(limit=2000))
        except Exception as exc:
            app.logger.warning("Could not fetch invoices from DB for suggestion: %s", exc)
            
    if storage is not None:
        try:
            all_invoices.extend(storage.list_invoices(limit=2000))
        except Exception as exc:
            app.logger.warning("Could not fetch invoices from storage for suggestion: %s", exc)
            
    for inv in all_invoices:
        inv_str = str(inv.get('invoice_number', '')).strip()
        if '2026-27' in inv_str or '26-27' in inv_str:
            m = re.match(r'^0*(\d+)', inv_str)
            if m:
                val = int(m.group(1))
                # Skip test bill 99
                if val != 99 and val > max_num:
                    max_num = val

    if max_num > 0:
        return f"{max_num + 1}{fy}"
    
    return f"1{fy}"


def format_sequential_invoice_number(base_inv: str, offset: int = 0) -> str:
    """Increment the sequence number in an invoice string (e.g. '036/2026-27' -> '037/2026-27')."""
    if not base_inv:
        return ''
    m = re.match(r'^(.*?)(\d+)(/.*)?$', base_inv)
    if m:
        prefix = m.group(1)
        digits = m.group(2)
        suffix = m.group(3) or ''
        seq_num = int(digits) + offset
        return f"{prefix}{str(seq_num).zfill(len(digits))}{suffix}"
    return base_inv


def format_date_for_invoice(date_str: str) -> str:
    """Convert YYYY-MM-DD to DD/MM/YYYY for invoice display."""
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        return dt.strftime('%d/%m/%Y')
    except ValueError:
        return date_str


def amount_in_words(amount: float) -> str:
    """Convert amount to words for invoice."""
    try:
        rupees = int(amount)
        paise = int(round((amount - rupees) * 100))
        
        if rupees == 0 and paise == 0:
            return "Zero Only"
        
        words = num2words(rupees, lang='en_IN').title()
        words = words.replace(',', '')
        
        if paise > 0:
            paise_words = num2words(paise, lang='en_IN').title()
            return f"Rupees {words} and {paise_words} Paise Only"
        
        return f"Rupees {words} Only"
    except Exception:
        return f"Rupees {amount} Only"


def safe_float(value: Any, default: float = 0.0) -> float:
    """Safely convert a value to float with a fallback."""
    try:
        if value is None or value == '':
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def round_half_up(value: float) -> int:
    """Round to nearest integer with .5 always rounding up (local app parity)."""
    return int(Decimal(str(value)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def _resolve_tax_scheme(
    tax_type: str,
    tax_rate_igst: float = 5.0,
    tax_rate_cgst: float = 2.5,
    tax_rate_sgst: float = 2.5,
) -> Dict[str, Any]:
    """Resolve a tax mode into display labels and numeric rates."""
    mode = (tax_type or 'IGST').strip().upper()
    if mode == 'PROFILE_DEFAULT':
        mode = 'IGST'

    default_igst_rate = 5.0
    default_cgst_rate = 2.5
    default_sgst_rate = 2.5

    if mode == 'CUSTOM_IGST':
        rate = max(0.0, safe_float(tax_rate_igst, default_igst_rate))
        return {
            'tax_type': mode,
            'display_type': 'IGST',
            'igst_rate': rate,
            'cgst_rate': 0.0,
            'sgst_rate': 0.0,
        }

    if mode == 'CUSTOM_CGST_SGST':
        cgst_rate = max(0.0, safe_float(tax_rate_cgst, default_cgst_rate))
        sgst_rate = max(0.0, safe_float(tax_rate_sgst, default_sgst_rate))
        return {
            'tax_type': mode,
            'display_type': 'CGST_SGST',
            'igst_rate': 0.0,
            'cgst_rate': cgst_rate,
            'sgst_rate': sgst_rate,
        }

    if mode == 'CGST_SGST':
        return {
            'tax_type': mode,
            'display_type': 'CGST_SGST',
            'igst_rate': 0.0,
            'cgst_rate': default_cgst_rate,
            'sgst_rate': default_sgst_rate,
        }

    return {
        'tax_type': 'IGST',
        'display_type': 'IGST',
        'igst_rate': default_igst_rate,
        'cgst_rate': 0.0,
        'sgst_rate': 0.0,
    }


def calculate_invoice_totals(
    items: List[Dict],
    tax_type: str = 'IGST',
    delivery_charge: float = 0.0,
    tax_rate_igst: float = 5.0,
    tax_rate_cgst: float = 2.5,
    tax_rate_sgst: float = 2.5,
) -> Dict:
    """Calculate invoice totals from items."""
    subtotal = sum(safe_float(item.get('quantity', 0), 0.0) * safe_float(item.get('rate', 0), 0.0) for item in items)
    delivery_charge = max(0.0, safe_float(delivery_charge, 0.0))
    taxable_amount = subtotal + delivery_charge
    tax_scheme = _resolve_tax_scheme(tax_type, tax_rate_igst, tax_rate_cgst, tax_rate_sgst)
    
    if tax_scheme['display_type'] == 'IGST':
        igst = round(taxable_amount * (tax_scheme['igst_rate'] / 100.0), 2)
        tax_amount = igst
        cgst = sgst = 0
    else:
        igst = 0
        cgst = round(taxable_amount * (tax_scheme['cgst_rate'] / 100.0), 2)
        sgst = round(taxable_amount * (tax_scheme['sgst_rate'] / 100.0), 2)
        tax_amount = cgst + sgst
    
    total_before_round = round(taxable_amount + tax_amount, 2)
    rounded_total = round(total_before_round)
    round_off = round(rounded_total - total_before_round, 2)
    
    return {
        'subtotal': round(subtotal, 2),
        'delivery_charge': round(delivery_charge, 2),
        'taxable_amount': round(taxable_amount, 2),
        'igst_amount': round(igst, 2),
        'cgst_amount': round(cgst, 2),
        'sgst_amount': round(sgst, 2),
        'tax_amount': round(tax_amount, 2),
        'tax_type': tax_scheme['tax_type'],
        'display_tax_type': tax_scheme['display_type'],
        'igst_rate': round(tax_scheme['igst_rate'], 2),
        'cgst_rate': round(tax_scheme['cgst_rate'], 2),
        'sgst_rate': round(tax_scheme['sgst_rate'], 2),
        'round_off_value': round(round_off, 2),
        'rounded_total': rounded_total,
        'amount_in_words': amount_in_words(rounded_total)
    }


def extract_transport_core(mode: str) -> str:
    """Extract core transport mode without prefix."""
    if not mode:
        return ''

    value = mode.strip()
    variants = [
        'mode of transport:',
        'mode of transport :',
        'mode of transports:',
        'mode of transports :',
        'transport:',
        'transport :',
    ]

    lowered = value.lower()
    for prefix in variants:
        if lowered.startswith(prefix):
            value = value[len(prefix):].strip(' -:')
            break

    return value.strip()


def normalize_transport_mode(mode: str) -> str:
    """Return canonical transport string used in Excel and storage."""
    core = extract_transport_core(mode)
    if not core:
        return ''
    return f"Mode of Transport: {core}"


def clean_address_lines(lines: Any) -> List[str]:
    """Strip redundant leading 'Buyer:' or 'Ship To:' header lines from address lists, phones/emails, and duplicate state labels."""
    if not lines:
        return []
    if isinstance(lines, str):
        lines = [line.strip() for line in re.split(r'[\r\n]+', lines) if line.strip()]
    cleaned = []
    has_canonical_state_code = any(re.search(r'(?i)\bstate\s*[-:]?\s*[A-Za-z\s]+[-:]?\s*code\s*[-:]?\s*\d{1,2}\b', str(l)) for l in lines)
    for line in lines:
        s = str(line).strip()
        if not s:
            continue
        if re.match(r'^(buyer|ship\s*to)\s*:\s*$', s, re.IGNORECASE):
            continue
        s = re.sub(r'^(buyer|ship\s*to)\s*:\s*', '', s, flags=re.IGNORECASE).strip()
        s = clean_line_phone_email(s)
        if not s:
            continue
        # Drop redundant standalone "State: <State>" or "State - <State>" lines when canonical state line exists or address line already contains the state
        if re.fullmatch(r'(?i)^\s*state\s*[:\-]\s*[A-Za-z\s]+$', s):
            if has_canonical_state_code:
                continue
            s_val = re.sub(r'(?i)^\s*state\s*[:\-]\s*', '', s).strip()
            if any(s_val.lower() in str(other).lower() for other in lines if other != line):
                continue
        cleaned.append(s)
    return cleaned


def _normalize_name(value: str) -> str:
    """Normalize names for robust profile matching."""
    if not value:
        return ''
    return re.sub(r'\s+', ' ', value.strip().lower())


def _match_buyer_profile(invoice: Dict, buyers: List[Dict]) -> Optional[Dict]:
    """Find the best buyer profile match for a stored invoice record."""
    if not buyers:
        return None

    profile_id = (invoice.get('buyer_profile_id') or '').strip()
    if profile_id:
        for buyer in buyers:
            if (buyer.get('profile_id') or '').strip() == profile_id:
                return buyer

    target_gstin = (invoice.get('buyer_gstin') or '').strip().upper()
    if target_gstin:
        for buyer in buyers:
            if (buyer.get('gstin') or '').strip().upper() == target_gstin:
                return buyer

    target_name = _normalize_name(invoice.get('buyer_name', ''))
    if target_name:
        for buyer in buyers:
            buyer_name = _normalize_name(buyer.get('buyer_name', ''))
            if buyer_name and (buyer_name == target_name or buyer_name in target_name or target_name in buyer_name):
                return buyer

    details = invoice.get('buyer_details', [])
    if isinstance(details, list):
        for line in details:
            line_norm = _normalize_name(str(line))
            if not line_norm or line_norm in {'buyer:', 'buyer :'}:
                continue
            for buyer in buyers:
                buyer_name = _normalize_name(buyer.get('buyer_name', ''))
                if buyer_name and (buyer_name == line_norm or buyer_name in line_norm or line_norm in buyer_name):
                    return buyer

    return None


# ===================== EXCEL GENERATION =====================

def generate_invoice_excel(invoice_data: Dict) -> bytes:
    """
    Generate an Excel invoice from the canonical 2026-27 template stored in Cloud Storage.
    The master template already contains the pixel-perfect signature natively embedded.
    Returns the Excel file as bytes.
    """
    template_bytes = None
    try:
        storage = get_cloud_storage()
        if storage:
            template_result = storage.download_template("invoice_template_2026_27.xlsx")
            if template_result:
                template_bytes, _ = template_result
    except Exception as e:
        app.logger.warning("Could not download template from Cloud Storage: %s", e)
    
    if not template_bytes:
        local_template = Path(__file__).parent / "invoice_template_2026_27.xlsx"
        if local_template.exists():
            template_bytes = local_template.read_bytes()
        else:
            raise Exception("Invoice template 'invoice_template_2026_27.xlsx' not found in Cloud Storage or locally")
    
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    sheet = wb.active

    # Force 1-page print settings for LibreOffice Calc
    sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = 'A1:I50'
    sheet.page_margins.left = 0.4
    sheet.page_margins.right = 0.4
    sheet.page_margins.top = 0.4
    sheet.page_margins.bottom = 0.4

    # --- SELLER DETAILS (ALWAYS SHAKAMBHARI ENTERPRISES) ---
    sheet['A4'] = "SELLER :"
    sheet['F4'] = "SHIP FROM:"
    sheet['A5'] = "M/S Shakambhari Enterprises."
    sheet['A6'] = "5th Floor, Flat 26B, Om Bhawan"
    sheet['A7'] = "54/5A, Strand Road, Jorabagan"
    sheet['A8'] = "Kolkata -700006, WB (19)"
    sheet['A9'] = "G.S.T No : 19ACZPN5725A1Z8"

    # --- INVOICE DETAILS ---
    sheet['A2'] = f"INVOICE No. {invoice_data.get('invoice_number', '')}"
    sheet['F2'] = f"Date : {invoice_data.get('invoice_date_display', '')}"
    
    ewaybill_num = str(invoice_data.get('ewaybill_number') or invoice_data.get('ewaybill_num') or '').strip()
    sheet['A3'] = f"Ewaybill No. {ewaybill_num}" if ewaybill_num else "Ewaybill No. "
        
    ewaybill_date = str(invoice_data.get('ewaybill_date_display') or invoice_data.get('ewaybill_date') or '').strip()
    if ewaybill_date and re.match(r'^\d{4}-\d{2}-\d{2}$', ewaybill_date):
        ewaybill_date = format_date_for_invoice(ewaybill_date)
    # If ewaybill no. is blank, ewaybill date must also be blank
    if ewaybill_num and ewaybill_date:
        sheet['F3'] = f"Ewaybill Date : {ewaybill_date}"
    else:
        sheet['F3'] = "Ewaybill Date : "

    # --- BUYER DETAILS ---
    buyer_details = clean_address_lines(invoice_data.get('buyer_details', []))
    for row_idx in range(13, 19):
        sheet[f'A{row_idx}'] = ''
        
    for i, detail in enumerate(buyer_details[:6]):
        sheet[f'A{13+i}'] = detail

    # --- SHIP FROM DETAILS ---
    ship_from = clean_address_lines(invoice_data.get('ship_from_details', []) or [])
    for row_idx in range(5, 10):
        sheet[f'F{row_idx}'] = ''
        
    for i, detail in enumerate(ship_from[:5]):
        sheet[f'F{5+i}'] = detail

    # --- SHIP TO DETAILS ---
    # Ship To always mirrors buyer details in the standard template
    raw_ship_to = invoice_data.get('ship_to_details') or invoice_data.get('buyer_details') or []
    ship_to = clean_address_lines(raw_ship_to)
    for row_idx in range(13, 19):
        sheet[f'F{row_idx}'] = ''
        
    for i, detail in enumerate(ship_to[:6]):
        sheet[f'F{13+i}'] = detail

    # --- TRANSPORT ---
    t_mode = normalize_transport_mode(invoice_data.get('transport_mode', ''))
    sheet['A20'] = t_mode if t_mode else ""

    # --- ITEMS ---
    items = invoice_data.get('items', [])
    first_item_row = 22
    thin_side = Side(border_style='thin', color='000000')

    # Reset all item rows (22 to 31) with clean borders (no horizontal dividers cutting empty rows)
    border_first_item = {
        'A': Border(top=thin_side, left=thin_side),
        'F': Border(top=thin_side, left=thin_side, right=thin_side),
        'G': Border(top=thin_side, left=thin_side, right=thin_side),
        'H': Border(top=thin_side, left=thin_side, right=thin_side),
        'I': Border(top=thin_side, left=thin_side, right=thin_side),
    }
    border_other_item = {
        'A': Border(left=thin_side),
        'F': Border(left=thin_side, right=thin_side),
        'G': Border(left=thin_side, right=thin_side),
        'H': Border(left=thin_side, right=thin_side),
        'I': Border(left=thin_side, right=thin_side),
    }

    for row_idx in range(22, 32):
        sheet[f'A{row_idx}'] = ''
        sheet[f'F{row_idx}'] = ''
        sheet[f'G{row_idx}'] = ''
        sheet[f'H{row_idx}'] = ''
        sheet[f'I{row_idx}'] = ''
        b_map = border_first_item if row_idx == 22 else border_other_item
        for col_letter in ['A', 'F', 'G', 'H', 'I']:
            sheet[f'{col_letter}{row_idx}'].border = b_map[col_letter]
        for col_letter in ['B', 'C', 'D', 'E']:
            sheet[f'{col_letter}{row_idx}'].border = Border(top=thin_side if row_idx == 22 else None)

    item_rows = []
    template_hsn = sheet['H22'].value or '76151030'
    base_font_name = sheet['A22'].font.name if sheet['A22'].font else 'Calibri'
    base_font_size = sheet['A22'].font.size if sheet['A22'].font else 10
    
    for idx, item in enumerate(items[:10]):
        row_num = first_item_row + idx
        item_rows.append(row_num)

        description = item.get('description', '').strip()
        # Ensure sequential numbering (1. ..., 2. ...) across both single and multi-item bills
        clean_desc = re.sub(r'^\d+\.\s*', '', description).strip()
        if clean_desc:
            description = f"{idx + 1}. {clean_desc}"

        hsn_value = item.get('hsn') or template_hsn
        quantity = safe_float(item.get('quantity', 0), 0.0)
        rate = safe_float(item.get('rate', 0), 0.0)

        sheet[f'A{row_num}'] = description
        sheet[f'A{row_num}'].alignment = Alignment(horizontal='left', vertical='center')
        sheet[f'A{row_num}'].font = Font(name=base_font_name, size=base_font_size, bold=False)

        sheet[f'F{row_num}'] = quantity
        sheet[f'F{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
        sheet[f'F{row_num}'].font = Font(name=base_font_name, size=base_font_size, bold=False)
        sheet[f'F{row_num}'].number_format = '0.000'

        sheet[f'G{row_num}'] = rate
        sheet[f'G{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
        sheet[f'G{row_num}'].font = Font(name=base_font_name, size=base_font_size, bold=False)
        sheet[f'G{row_num}'].number_format = '0.00'

        sheet[f'H{row_num}'] = hsn_value
        sheet[f'H{row_num}'].alignment = Alignment(horizontal='center', vertical='center')
        sheet[f'H{row_num}'].font = Font(name=base_font_name, size=base_font_size, bold=True)

        sheet[f'I{row_num}'] = f'=F{row_num}*G{row_num}'
        sheet[f'I{row_num}'].alignment = Alignment(horizontal='right', vertical='center')
        sheet[f'I{row_num}'].font = Font(name=base_font_name, size=base_font_size, bold=False)
        sheet[f'I{row_num}'].number_format = '0.00'

    # --- SUBTOTALS & TAXES ---
    if len(item_rows) == 1:
        subtotal_formula = f'=I{item_rows[0]}'
        qty_formula = f'=F{item_rows[0]}'
    elif len(item_rows) > 1:
        subtotal_formula = f'=SUM(I{item_rows[0]}:I{item_rows[-1]})'
        qty_formula = f'=SUM(F{item_rows[0]}:F{item_rows[-1]})'
    else:
        subtotal_formula = '=0'
        qty_formula = '=0'
        
    sheet['F32'] = qty_formula
    sheet['F32'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['F32'].font = Font(name=base_font_name, size=base_font_size, bold=False)
    sheet['F32'].number_format = '0.000'

    sheet['I32'] = subtotal_formula
    sheet['I32'].alignment = Alignment(horizontal='right', vertical='center')
    sheet['I32'].font = Font(name=base_font_name, size=base_font_size, bold=False)
    sheet['I32'].number_format = '0.00'

    # Closing border on subtotal row 32
    for col_c in ['F', 'G', 'H', 'I']:
        sheet[f'{col_c}32'].border = Border(top=thin_side, bottom=thin_side, left=thin_side, right=thin_side)
    
    delivery_charge = max(0.0, safe_float(invoice_data.get('delivery_charge', 0), 0.0))
    sheet['C33'] = 'Delivery Charge'
    sheet['C33'].alignment = Alignment(horizontal='left', vertical='center')
    sheet['I33'] = delivery_charge
    sheet['I33'].alignment = Alignment(horizontal='right', vertical='center')
    sheet['I33'].number_format = '0.00'

    tax_type = invoice_data.get('tax_type', 'IGST')
    tax_scheme = _resolve_tax_scheme(
        tax_type,
        invoice_data.get('tax_rate_igst', 5.0),
        invoice_data.get('tax_rate_cgst', 2.5),
        invoice_data.get('tax_rate_sgst', 2.5),
    )
    
    tax_base_formula = '(I32+I33)'
    
    if tax_scheme['display_type'] == 'IGST':
        sheet['C34'] = 'G.S.T SALES I.G.S.T @'
        sheet['C34'].alignment = Alignment(horizontal='left', vertical='center')
        sheet['E34'] = f"{tax_scheme['igst_rate']:.2f}%"
        sheet['E34'].alignment = Alignment(horizontal='right', vertical='center')
        sheet['I34'] = f'=ROUND({tax_base_formula}*{tax_scheme["igst_rate"]}/100, 2)'
        sheet['I34'].alignment = Alignment(horizontal='right', vertical='center')
        sheet['I34'].number_format = '0.00'
        # Clear Row 35 completely to eliminate overlapping text artifacts in LibreOffice
        sheet['C35'] = ''
        sheet['E35'] = ''
        sheet['I35'] = ''
    else:
        sheet['C34'] = 'G.S.T SALES C.G.S.T @'
        sheet['C34'].alignment = Alignment(horizontal='left', vertical='center')
        sheet['E34'] = f"{tax_scheme['cgst_rate']:.2f}%"
        sheet['E34'].alignment = Alignment(horizontal='right', vertical='center')
        sheet['I34'] = f'=ROUND({tax_base_formula}*{tax_scheme["cgst_rate"]}/100, 2)'
        sheet['I34'].alignment = Alignment(horizontal='right', vertical='center')
        sheet['I34'].number_format = '0.00'
        sheet['C35'] = 'G.S.T SALES S.G.S.T @'
        sheet['C35'].alignment = Alignment(horizontal='left', vertical='center')
        sheet['E35'] = f"{tax_scheme['sgst_rate']:.2f}%"
        sheet['E35'].alignment = Alignment(horizontal='right', vertical='center')
        sheet['I35'] = f'=ROUND({tax_base_formula}*{tax_scheme["sgst_rate"]}/100, 2)'
        sheet['I35'].alignment = Alignment(horizontal='right', vertical='center')
        sheet['I35'].number_format = '0.00'

    # Calculation and round off
    sheet['I41'] = '=SUM(I32:I35)'
    sheet['I37'] = '=ROUND(I41,0)-I41'
    sheet['I37'].alignment = Alignment(horizontal='right', vertical='center')
    sheet['I37'].number_format = '0.00'

    sheet['I38'] = '=ROUND(I41,0)'
    sheet['I38'].alignment = Alignment(horizontal='right', vertical='center')
    sheet['I38'].number_format = '0.00'
    sheet['I38'].font = Font(name=base_font_name, size=base_font_size, bold=True)
    
    # Ensure all tax and calculation rows remain visible with fixed spacing
    sheet.row_dimensions[33].hidden = False
    sheet.row_dimensions[34].hidden = False
    sheet.row_dimensions[35].hidden = False
    sheet.row_dimensions[36].hidden = False
    sheet.row_dimensions[37].hidden = False
    sheet.row_dimensions[38].hidden = False
    sheet.row_dimensions[41].hidden = True
    sheet.row_dimensions[41].height = 0.6
    
    # --- AMOUNT IN WORDS ---
    subtotal = sum(safe_float(item.get('quantity', 0), 0.0) * safe_float(item.get('rate', 0), 0.0) for item in items)
    tax_base_value = subtotal + delivery_charge

    if tax_scheme['display_type'] == 'IGST':
        tax_amount = tax_base_value * (tax_scheme['igst_rate'] / 100.0)
    else:
        tax_amount = tax_base_value * ((tax_scheme['cgst_rate'] + tax_scheme['sgst_rate']) / 100.0)
        
    total_before_round = tax_base_value + tax_amount
    rounded_total = round_half_up(total_before_round)
    
    if rounded_total > 0:
        words = num2words(int(rounded_total), lang='en_IN').replace('-', ' ').replace(',', ' ').title()
        amount_words = f"AMOUNT : {words} Only"
    else:
        amount_words = 'AMOUNT : Zero Only'
    sheet['A40'] = amount_words
    
    # --- BANK & SIGNATORY DETAILS (ALWAYS SHAKAMBHARI ENTERPRISES) ---
    sheet['A42'] = "Bank Details:"
    sheet['A43'] = "ICICI Bank"
    sheet['A44'] = "Salt lake Sector 2 Branch"
    sheet['A45'] = "A/C No.: 004205008313 "
    sheet['A46'] = "IFSC No.: ICIC0000042 "
    sheet['G45'] = "For Shakambhari Enterprises"
    sheet['G48'] = "Authorised Signatory"
    # Handle signature image
    is_signed = bool(invoice_data.get('digitally_signed', True))
    if not is_signed:
        if hasattr(sheet, '_images'):
            sheet._images = []
    else:
        has_img = hasattr(sheet, '_images') and len(sheet._images) > 0
        if not has_img:
            sig_path = Path(__file__).parent / 'static' / 'signature.png'
            if not sig_path.exists():
                sig_path = Path('static') / 'signature.png'
            if sig_path.exists():
                try:
                    from openpyxl.drawing.image import Image as OpenpyxlImage
                    sig_img = OpenpyxlImage(str(sig_path))
                    sig_img.height = 55
                    sig_img.width = int(55 * (812 / 395))
                    sheet.add_image(sig_img, 'G45')
                except Exception as img_err:
                    app.logger.warning("Could not add fallback signature image: %s", img_err)

    sheet['A50'] = "Subject to Kolkata Jurisdiction."

    # Add layout marker
    sheet['Z1'] = 'v2'
    sheet.column_dimensions['Z'].hidden = True


    try:
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.read()
    finally:
        wb.close()


def generate_pdf_from_excel(excel_bytes: bytes, excel_filename: str) -> Optional[bytes]:
    """
    Generate PDF directly from XLSX using LibreOffice headless.
    This provides visual output closest to the spreadsheet layout.
    """
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        app.logger.error('LibreOffice (soffice) is not installed in this runtime.')
        return None

    safe_xlsx_name = _safe_filename(excel_filename or 'invoice.xlsx', '.xlsx')

    with tempfile.TemporaryDirectory(prefix='xlsx_to_pdf_') as tmpdir:
        xlsx_path = os.path.join(tmpdir, safe_xlsx_name)
        with open(xlsx_path, 'wb') as f:
            f.write(excel_bytes)

        cmd = [
            soffice,
            '--headless',
            '--nologo',
            '--nolockcheck',
            '--nodefault',
            '--nofirststartwizard',
            '--convert-to', 'pdf:calc_pdf_Export',
            '--outdir', tmpdir,
            xlsx_path,
        ]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
        except Exception as exc:
            app.logger.exception('Failed to run LibreOffice conversion: %s', exc)
            return None

        if result.returncode != 0:
            app.logger.error(
                'LibreOffice conversion failed (code=%s): stdout=%s stderr=%s',
                result.returncode,
                (result.stdout or '').strip(),
                (result.stderr or '').strip(),
            )
            return None

        expected_pdf = os.path.splitext(xlsx_path)[0] + '.pdf'
        pdf_path = expected_pdf if os.path.exists(expected_pdf) else ''
        if not pdf_path:
            # Fallback in case LibreOffice produces a different basename.
            for name in os.listdir(tmpdir):
                if name.lower().endswith('.pdf'):
                    pdf_path = os.path.join(tmpdir, name)
                    break

        if not pdf_path or not os.path.exists(pdf_path):
            app.logger.error('LibreOffice conversion completed but no PDF output was found.')
            return None

        with open(pdf_path, 'rb') as f:
            return f.read()


# ===================== PWA & STATIC ASSET ROUTES =====================

@app.route('/manifest.json')
def manifest_json():
    """Serve PWA Web App Manifest."""
    static_dir = os.path.join(os.path.dirname(__file__), 'static')
    return send_from_directory(static_dir, 'manifest.json', mimetype='application/manifest+json')


@app.route('/sw.js')
def service_worker():
    """Serve PWA Service Worker with root scope permissions."""
    static_dir = os.path.join(os.path.dirname(__file__), 'static')
    resp = make_response(send_from_directory(static_dir, 'sw.js', mimetype='application/javascript'))
    resp.headers['Service-Worker-Allowed'] = '/'
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return resp


@app.route('/favicon.ico')
def favicon():
    """Serve favicon."""
    static_dir = os.path.join(os.path.dirname(__file__), 'static')
    return send_from_directory(static_dir, 'favicon.ico', mimetype='image/x-icon')


# ===================== ROUTE HANDLERS =====================

@app.route('/')
def index():
    """Main invoice generation page."""
    today_date = get_ist_today_str()
    suggestion = suggest_next_invoice_number()
    db = get_sheets_db_or_none()
    bucket = os.environ.get('GCS_BUCKET_NAME', '')
    project = os.environ.get('GOOGLE_CLOUD_PROJECT', '')
    bucket_base = f"https://console.cloud.google.com/storage/browser/{bucket}"
    project_suffix = f"?project={project}" if project else ''

    storage = None
    try:
        storage = get_cloud_storage()
    except Exception as exc:
        app.logger.warning("Could not initialize Cloud Storage in index: %s", exc)

    if db is None:
        flash("Sheets configuration is missing, so saved profiles and invoice history are unavailable.", "warning")
        return render_template(
            'index.html',
            buyer_profiles=[],
            transport_modes=[],
            dispatch_addresses=list(getattr(GoogleSheetsDB, 'DEFAULT_DISPATCH_ADDRESSES', [])),
            today_date=today_date,
            suggested_invoice_number=suggestion,
            recent_invoices=[],
            preload_invoice=None,
            open_records=(request.args.get('open_records') == '1'),
            app_settings=get_settings(storage),
            bucket_console_url=f"{bucket_base}{project_suffix}" if bucket else '',
            invoices_folder_url=f"{bucket_base}/invoices/{project_suffix}" if bucket else '',
            pdfs_folder_url=f"{bucket_base}/pdfs/{project_suffix}" if bucket else '',
        )

    try:
        buyer_profiles = db.get_all_buyers()
    except Exception as exc:
        app.logger.warning("Could not fetch buyers: %s", exc)
        buyer_profiles = []

    try:
        transport_modes = db.get_all_transport_modes()
    except Exception as exc:
        app.logger.warning("Could not fetch transport modes: %s", exc)
        transport_modes = []

    buyer_profiles.sort(key=lambda p: p.get('buyer_name', '').lower())

    try:
        recent_invoices = _build_invoice_rows(db.get_all_invoices(limit=2000))
    except Exception as exc:
        app.logger.warning("Could not fetch recent invoices: %s", exc)
        recent_invoices = []

    invoice_transport_modes = [inv.get('transport_mode', '') for inv in recent_invoices if inv.get('transport_mode')]
    combined_transport_modes = [m for m in (transport_modes + invoice_transport_modes) if m]
    transport_cores = list(set(extract_transport_core(m) for m in combined_transport_modes if extract_transport_core(m)))
    transport_cores.sort()
    if storage:
        try:
            storage_rows = storage.list_invoices(limit=2000)
            recent_invoices = _merge_with_storage_rows(recent_invoices, storage_rows)
        except Exception as exc:
            app.logger.warning('Could not list storage invoices for modal merge: %s', exc)

    # Deduplicate invoices by unique file or distinct bill signature so multiple bills sharing
    # the same sequential number (e.g. 61/2026-27) are preserved individually in the past records list.
    deduped = {}
    for inv in recent_invoices:
        inv_num = str(inv.get('invoice_number', '')).strip()
        if not inv_num:
            continue
        fn = str(inv.get('filename') or '').strip().lower()
        if fn:
            unique_key = fn
        else:
            buyer = str(inv.get('buyer_name') or '').strip().lower()
            date = str(inv.get('invoice_date') or '').strip().lower()
            amt = str(inv.get('total_amount') or '').strip().lower()
            unique_key = f"{inv_num.lower()}_{buyer}_{date}_{amt}"

        if unique_key not in deduped:
            deduped[unique_key] = inv
        else:
            prev = deduped[unique_key]
            if (len(inv.get('items') or []) > len(prev.get('items') or [])) or (inv.get('file_url') and not prev.get('file_url')):
                deduped[unique_key] = inv

    recent_invoices = list(deduped.values())
    recent_invoices.sort(key=_invoice_sort_key, reverse=True)


    load_invoice_number = request.args.get('load', '').strip()
    preload_invoice = None
    if load_invoice_number:
        if db is not None:
            try:
                preload_invoice = db.get_invoice(load_invoice_number)
            except Exception as exc:
                app.logger.warning("Could not fetch preload invoice %s from Sheets: %s", load_invoice_number, exc)
        if not preload_invoice and storage:
            pattern = _safe_filename(load_invoice_number.replace('/', '-'), '')
            try:
                blobs = list(storage.client.list_blobs(storage.bucket_name, prefix=f"{storage.INVOICES_FOLDER}Invoice_{pattern}"))
                if blobs:
                    name = blobs[0].name.replace(storage.INVOICES_FOLDER, '')
                    file_bytes = storage.download_invoice_xlsx(name)
                    if file_bytes:
                        preload_invoice = _extract_invoice_data_from_xlsx_bytes(file_bytes, name)
            except Exception as exc:
                app.logger.warning("Could not fetch preload invoice %s from GCS: %s", load_invoice_number, exc)
        if preload_invoice:
            preload_invoice = _enrich_invoice_for_frontend(preload_invoice)

    try:
        dispatch_addresses = db.get_all_dispatch_addresses() if db else list(getattr(GoogleSheetsDB, 'DEFAULT_DISPATCH_ADDRESSES', []))
    except Exception as exc:
        app.logger.warning("Could not fetch dispatch addresses: %s", exc)
        dispatch_addresses = list(getattr(GoogleSheetsDB, 'DEFAULT_DISPATCH_ADDRESSES', []))

    for inv in recent_invoices:
        sf = inv.get('ship_from_details')
        if sf:
            sf_str = '\n'.join(sf) if isinstance(sf, list) else str(sf).strip()
            if sf_str and sf_str not in dispatch_addresses:
                dispatch_addresses.append(sf_str)

    return render_template('index.html',
                          buyer_profiles=buyer_profiles,
                          transport_modes=transport_cores,
                          dispatch_addresses=dispatch_addresses,
                          today_date=today_date,
                          suggested_invoice_number=suggestion,
                          recent_invoices=recent_invoices,
                          preload_invoice=preload_invoice,
                          open_records=(request.args.get('open_records') == '1'),
                          app_settings=sanitize_settings_for_client(get_settings(storage)),
                          bucket_console_url=f"{bucket_base}{project_suffix}" if bucket else '',
                          invoices_folder_url=f"{bucket_base}/invoices/{project_suffix}" if bucket else '',
                          pdfs_folder_url=f"{bucket_base}/pdfs/{project_suffix}" if bucket else '')


@app.route('/dashboard')
def dashboard():
    """Dashboard is merged into index modal; keep this route as a compatibility redirect."""
    return redirect(url_for('index', open_records='1'))


@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    """Generate an invoice from form data."""
    try:
        db = get_sheets_db()
        storage = get_cloud_storage()
        
        # Get form data
        buyer_profile_id = request.form.get('buyer_profile_id')
        if not buyer_profile_id:
            flash("Please select a buyer profile.", "error")
            return redirect(url_for('index'))
        
        invoice_number = request.form.get('invoice_number', '').strip()
        invoice_date_str = request.form.get('invoice_date', '')
        ewaybill_number = request.form.get('ewaybill_number', '').strip()
        ewaybill_date_str = request.form.get('ewaybill_date', '').strip()
        transport_mode_input = request.form.get('transport_mode_input', request.form.get('transport_mode', '')).strip()
        transport_mode = normalize_transport_mode(transport_mode_input)
        delivery_charge = safe_float(request.form.get('delivery_charge', '0').strip(), 0.0)
        if delivery_charge < 0:
            flash("Delivery charge cannot be negative.", "error")
            return redirect(url_for('index'))
        tax_type_override = request.form.get('tax_type_override', 'PROFILE_DEFAULT')
        tax_rate_igst = safe_float(request.form.get('tax_rate_igst', '5').strip(), 5.0)
        tax_rate_cgst = safe_float(request.form.get('tax_rate_cgst', '2.5').strip(), 2.5)
        tax_rate_sgst = safe_float(request.form.get('tax_rate_sgst', '2.5').strip(), 2.5)
        ship_from_enabled = bool(request.form.get('ship_from_enabled'))
        ship_from_text = request.form.get('ship_from', '').strip() if ship_from_enabled else ''
        digitally_signed = bool(request.form.get('digitally_signed', True))
        
        # Get buyer profile
        buyer = db.get_buyer(buyer_profile_id)
        if not buyer:
            flash("Buyer profile not found.", "error")
            return redirect(url_for('index'))
        
        # Determine tax type
        if tax_type_override == 'PROFILE_DEFAULT':
            tax_type = buyer.get('default_tax_type', 'IGST')
        else:
            tax_type = tax_type_override
        
        # Parse items
        descriptions = request.form.getlist('item_description[]')
        bags = request.form.getlist('item_bags[]')
        units = request.form.getlist('item_unit[]')
        item_hsns = request.form.getlist('item_hsn[]')
        quantities = request.form.getlist('item_quantity[]')
        rates = request.form.getlist('item_rate[]')
        
        items = []
        for i in range(len(descriptions)):
            try:
                qty = float(quantities[i]) if quantities[i] else 0
                rate = float(rates[i]) if rates[i] else 0
                
                if qty > 0 or rate > 0:
                    desc = descriptions[i].strip()
                    raw_bag = bags[i].strip() if i < len(bags) and bags[i] else ''
                    unit_val = units[i].strip() if i < len(units) and units[i].strip() else 'Bags'
                    
                    # Clean bag value: strip any unit text so only the count remains
                    clean_bag = re.sub(r'(?i)\s*(?:bags?|bundles?|boxes?|pkts?|cartons?)', '', raw_bag).strip()
                    
                    # Prevent duplication by stripping any existing packaging suffix from description
                    desc = re.sub(r'\s*\(\s*[\d\.]+\s*(?:[A-Za-z]+|\s*)+\)', '', desc, flags=re.IGNORECASE).strip()
                    
                    if clean_bag:
                        desc += f" ({clean_bag} {unit_val})"
                    
                    hsn = item_hsns[i].strip() if i < len(item_hsns) else ''
                    items.append({
                        'description': desc,
                        'bags': clean_bag,
                        'unit': unit_val,
                        'hsn': hsn,
                        'quantity': qty,
                        'rate': rate,
                        'amount': qty * rate
                    })
            except (ValueError, IndexError):
                continue
        
        if not items:
            flash("Please add at least one item.", "error")
            return redirect(url_for('index'))
        
        # Calculate totals
        totals = calculate_invoice_totals(
            items,
            tax_type,
            delivery_charge=delivery_charge,
            tax_rate_igst=tax_rate_igst,
            tax_rate_cgst=tax_rate_cgst,
            tax_rate_sgst=tax_rate_sgst,
        )
        
        def _normalized_multiline_lines(value):
            return [line.strip() for line in re.split(r'[\r\n]+', value) if line.strip()]

        ship_from_lines = _normalized_multiline_lines(ship_from_text)
        # Ship To always mirrors buyer details in the canonical template
        ship_to_lines = buyer.get('buyer_details', [])
        ship_to_enabled = True

        # Prepare invoice data
        invoice_data = {
            'invoice_number': invoice_number,
            'invoice_date': invoice_date_str,
            'invoice_date_display': format_date_for_invoice(invoice_date_str),
            'buyer_name': buyer['buyer_name'],
            'buyer_gstin': buyer.get('gstin', ''),
            'buyer_details': buyer.get('buyer_details', []),
            'ewaybill_number': ewaybill_number,
            'ewaybill_date': ewaybill_date_str,
            'ewaybill_date_display': format_date_for_invoice(ewaybill_date_str) if ewaybill_date_str else '',
            'ship_from_details': ship_from_lines,
            'ship_to_enabled': ship_to_enabled,
            'ship_to_details': ship_to_lines,
            'digitally_signed': digitally_signed,
            'items': items,
            'transport_mode': transport_mode,
            'delivery_charge': delivery_charge,
            'tax_type': tax_type,
            'display_tax_type': totals['display_tax_type'],
            'tax_rate_igst': totals['igst_rate'],
            'tax_rate_cgst': totals['cgst_rate'],
            'tax_rate_sgst': totals['sgst_rate'],
            **totals
        }
        
        # Generate Excel file
        excel_bytes = generate_invoice_excel(invoice_data)
        
        # Generate filename
        safe_buyer = ''.join(c if c.isalnum() else '_' for c in buyer['buyer_name'][:20])
        filename = f"Invoice_{invoice_number.replace('/', '-')}_{invoice_date_str}_{safe_buyer}.xlsx"
        
        # Upload to Cloud Storage
        xlsx_url = storage.upload_invoice_xlsx(excel_bytes, filename)
        
        # Generate and upload PDF
        pdf_url = ''
        pdf_filename = filename.replace('.xlsx', '.pdf')
        
        pdf_bytes = generate_pdf_from_excel(excel_bytes, filename)
        if pdf_bytes:
            pdf_url = storage.upload_invoice_pdf(pdf_bytes, pdf_filename)
        else:
            app.logger.error('Skipping PDF upload because XLSX-to-PDF conversion failed for %s', filename)
        
        # Save invoice record to Google Sheets
        invoice_record = {
            'invoice_number': invoice_number,
            'invoice_date': invoice_date_str,
            'buyer_name': buyer['buyer_name'],
            'buyer_gstin': buyer.get('gstin', ''),
            'items': items,
            'subtotal': totals['subtotal'],
            'tax_type': tax_type,
            'display_tax_type': totals['display_tax_type'],
            'tax_rate_igst': totals['igst_rate'],
            'tax_rate_cgst': totals['cgst_rate'],
            'tax_rate_sgst': totals['sgst_rate'],
            'tax_amount': totals['tax_amount'],
            'total_amount': totals['rounded_total'],
            'transport_mode': transport_mode,
            'file_url': xlsx_url,
            'pdf_url': pdf_url,
            'ewaybill_number': ewaybill_number,
            'ewaybill_date': ewaybill_date_str,
            'ship_from_details': ship_from_lines,
            'ship_to_enabled': ship_to_enabled,
            'ship_to_details': ship_to_lines,
            'digitally_signed': digitally_signed,
            'delivery_charge': delivery_charge,
        }
        # Save invoice record to Google Sheets (non-blocking if Sheets API is rate-limited)
        try:
            db.save_invoice(invoice_record)
            if transport_mode:
                db.add_transport_mode(transport_mode)
            if ship_from_text:
                db.add_dispatch_address(ship_from_text)
        except Exception as sheet_err:
            app.logger.warning("Could not sync invoice metadata to Google Sheets (rate limit/network): %s", sheet_err)
        
        flash(f"Invoice {invoice_number} generated successfully!", "success")

        excel_download_url = url_for('download_xlsx', filename=filename)
        pdf_download_url = url_for('download_pdf', filename=pdf_filename) if pdf_url else ''
        
        return render_template('success.html',
                              filename=filename,
                              invoice_number=invoice_number,
                      excel_url=excel_download_url,
                      pdf_url=pdf_download_url,
                      is_pdf=bool(pdf_url))
        
    except Exception as e:
        flash(f"Error generating invoice: {str(e)}", "error")
        import traceback
        traceback.print_exc()
        return redirect(url_for('index'))


def _resolve_and_fetch_invoice_file(target: str, ext: str) -> Tuple[Optional[bytes], str]:
    """
    Find or generate invoice file (XLSX or PDF) by exact filename or invoice number.
    Returns (file_bytes, filename) or (None, '').
    """
    storage = get_cloud_storage()
    if not storage:
        return None, ''

    ext = ext.lower()
    if not ext.startswith('.'):
        ext = '.' + ext

    # 1. Direct match if target is an exact filename
    cleaned_target = os.path.basename(target.strip())
    if cleaned_target.lower().endswith(ext):
        file_bytes = storage.download_invoice_xlsx(cleaned_target) if ext == '.xlsx' else storage.download_invoice_pdf(cleaned_target)
        if file_bytes:
            return file_bytes, cleaned_target

    # 2. Extract invoice number or search pattern
    num_m = re.search(r'(\d+)', target)
    num_val = int(num_m.group(1)) if num_m else None
    pat = target.replace('/', '-').replace(' ', '_').strip()

    matching_xlsx_blob = None
    matching_pdf_blob = None

    try:
        blobs = list(storage.client.list_blobs(storage.bucket_name, prefix=storage.INVOICES_FOLDER))
        for b in blobs:
            bname = b.name.replace(storage.INVOICES_FOLDER, '')
            if ext == '.pdf' and bname.lower().endswith('.pdf'):
                if (pat and pat.lower() in bname.lower()) or (num_val is not None and re.search(rf'Invoice_0*{num_val}[-_]', bname)):
                    matching_pdf_blob = bname
                    break
            if bname.lower().endswith('.xlsx'):
                if (pat and pat.lower() in bname.lower()) or (num_val is not None and re.search(rf'Invoice_0*{num_val}[-_]', bname)):
                    matching_xlsx_blob = bname
                    if ext == '.xlsx':
                        break
    except Exception as exc:
        app.logger.warning("Error searching GCS for invoice %s: %s", target, exc)

    if ext == '.pdf':
        if matching_pdf_blob:
            b_bytes = storage.download_invoice_pdf(matching_pdf_blob)
            if b_bytes:
                return b_bytes, matching_pdf_blob
        # If no PDF blob, but matching XLSX exists, convert on the fly!
        if matching_xlsx_blob:
            xlsx_bytes = storage.download_invoice_xlsx(matching_xlsx_blob)
            if xlsx_bytes:
                pdf_bytes = generate_pdf_from_excel(xlsx_bytes, matching_xlsx_blob)
                if pdf_bytes:
                    pdf_name = os.path.splitext(matching_xlsx_blob)[0] + '.pdf'
                    try:
                        storage.upload_invoice_pdf(pdf_bytes, pdf_name)
                    except Exception:
                        pass
                    return pdf_bytes, pdf_name
    elif ext == '.xlsx':
        if matching_xlsx_blob:
            b_bytes = storage.download_invoice_xlsx(matching_xlsx_blob)
            if b_bytes:
                return b_bytes, matching_xlsx_blob

    # 3. Fallback: Lookup invoice in Sheets DB and generate on the fly
    db = get_sheets_db_or_none()
    if db:
        try:
            inv = db.get_invoice(target)
            if inv:
                xlsx_bytes = generate_invoice_excel(inv)
                safe_inv_num = inv.get('invoice_number', 'invoice').replace('/', '-')
                buyer_part = _safe_filename(inv.get('buyer_name', ''), '').replace('.xlsx', '')
                xlsx_name = f"Invoice_{safe_inv_num}_{buyer_part}.xlsx"
                try:
                    storage.upload_invoice_xlsx(xlsx_bytes, xlsx_name)
                except Exception:
                    pass
                if ext == '.xlsx':
                    return xlsx_bytes, xlsx_name
                elif ext == '.pdf':
                    pdf_bytes = generate_pdf_from_excel(xlsx_bytes, xlsx_name)
                    if pdf_bytes:
                        pdf_name = os.path.splitext(xlsx_name)[0] + '.pdf'
                        try:
                            storage.upload_invoice_pdf(pdf_bytes, pdf_name)
                        except Exception:
                            pass
                        return pdf_bytes, pdf_name
        except Exception as exc:
            app.logger.warning("Sheets fallback generation failed for %s: %s", target, exc)

    return None, ''


@app.route('/download/xlsx/<path:filename>')
@app.route('/download_xlsx/<path:filename>')
def download_xlsx(filename):
    """Download an Excel invoice by filename or invoice number."""
    file_bytes, safe_name = _resolve_and_fetch_invoice_file(filename, '.xlsx')
    if not file_bytes:
        flash("Excel invoice file could not be found.", "error")
        return redirect(url_for('index'))
    return Response(
        file_bytes,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{safe_name}"'}
    )


@app.route('/download/pdf/<path:filename>')
@app.route('/download_pdf/<path:filename>')
def download_pdf(filename):
    """Download a PDF invoice by filename or invoice number."""
    file_bytes, safe_name = _resolve_and_fetch_invoice_file(filename, '.pdf')
    if not file_bytes:
        flash("PDF invoice file could not be found.", "error")
        return redirect(url_for('index'))
    return Response(
        file_bytes,
        mimetype='application/pdf',
        headers={'Content-Disposition': f'attachment; filename="{safe_name}"'}
    )


# ===================== BUYER PROFILE ROUTES =====================

@app.route('/profiles')
def list_profiles():
    """List all buyer profiles."""
    try:
        db = get_sheets_db_or_none()
        if db is None:
            raise ValueError("Sheets configuration is missing")
        profiles = db.get_all_buyers()
        profiles.sort(key=lambda p: p.get('buyer_name', '').lower())
    except Exception:
        app.logger.exception("Failed to load buyer profiles")
        flash(
            "Buyer profiles could not be loaded right now. Check your Sheets configuration.",
            "error",
        )
        profiles = []
    return render_template('list_profiles.html', profiles=profiles)


@app.route('/profile', methods=['GET', 'POST'])
@app.route('/profile/<profile_id>', methods=['GET', 'POST'])
def manage_profile(profile_id=None):
    """Create or edit a buyer profile."""
    db = get_sheets_db_or_none()
    
    is_new_profile = profile_id is None
    profile = None

    if db is None:
        flash("Set SPREADSHEET_ID and Google credentials before managing profiles.", "error")
        profile = {
            'buyer_name': '',
            'buyer_details': [],
            'buyer_details_textarea': '',
            'gstin': '',
            'default_tax_type': 'IGST',
            'profile_id': profile_id or ''
        }
        return render_template('profile_form.html', profile=profile, is_new_profile=is_new_profile)
    
    if profile_id:
        profile = db.get_buyer(profile_id)
        if not profile:
            flash("Profile not found.", "error")
            return redirect(url_for('list_profiles'))

    if profile and request.method == 'GET':
        profile['buyer_details_textarea'] = _details_for_edit_form(profile)
    
    if request.method == 'POST':
        buyer_name = request.form.get('buyer_name', '').strip()
        buyer_details_str = request.form.get('buyer_details_textarea', '')
        fallback_gstin = (profile or {}).get('gstin', '')
        buyer_details, gstin = _normalize_buyer_details_for_storage(buyer_name, buyer_details_str, fallback_gstin)
        default_tax_type = request.form.get('default_tax_type', 'IGST')
        
        if not buyer_name:
            flash("Buyer Name is required.", "error")
            profile_data = {
                'buyer_name': buyer_name,
                'buyer_details_textarea': buyer_details_str,
                'buyer_details': buyer_details,
                'gstin': gstin,
                'default_tax_type': default_tax_type,
                'profile_id': profile_id or ''
            }
            return render_template('profile_form.html', profile=profile_data,
                                 is_new_profile=is_new_profile)

        if not buyer_details_str.strip():
            flash("Buyer address details are required.", "error")
            profile_data = {
                'buyer_name': buyer_name,
                'buyer_details_textarea': buyer_details_str,
                'buyer_details': buyer_details,
                'gstin': gstin,
                'default_tax_type': default_tax_type,
                'profile_id': profile_id or gstin or ''
            }
            return render_template('profile_form.html', profile=profile_data,
                                 is_new_profile=is_new_profile)
        
        if is_new_profile:
            # Generate profile ID
            new_profile_id = profile_id or gstin or f"{buyer_name.replace(' ', '_')}_{uuid.uuid4().hex[:8]}"
        else:
            new_profile_id = profile_id
        
        raw_p = {
            'profile_id': new_profile_id,
            'buyer_name': buyer_name,
            'buyer_details': buyer_details,
            'gstin': gstin,
            'default_tax_type': default_tax_type
        }
        profile_data = normalize_buyer_profile(raw_p)
        
        if db.save_buyer(profile_data):
            flash(f"Profile '{buyer_name}' saved successfully!", "success")
            return redirect(url_for('list_profiles'))
        else:
            flash("Error saving profile.", "error")
    
    # GET request
    if profile:
        profile['buyer_details_textarea'] = '\n'.join(profile.get('buyer_details', []))
    else:
        profile = {
            'buyer_name': '',
            'buyer_details': [],
            'buyer_details_textarea': '',
            'gstin': '',
            'default_tax_type': 'IGST',
            'profile_id': ''
        }
    
    return render_template('profile_form.html', profile=profile, is_new_profile=is_new_profile)


@app.route('/profile/<profile_id>/delete', methods=['POST'])
def delete_profile(profile_id):
    """Delete a buyer profile."""
    db = get_sheets_db_or_none()
    if db is None:
        flash("Set SPREADSHEET_ID and Google credentials before deleting profiles.", "error")
        return redirect(url_for('list_profiles'))
    
    if db.delete_buyer(profile_id):
        flash("Profile deleted successfully.", "success")
    else:
        flash("Profile not found.", "error")
    
    return redirect(url_for('list_profiles'))


# ===================== API ROUTES =====================

@app.route('/api/calculate', methods=['POST'])
def calculate_preview():
    """API endpoint to calculate invoice preview."""
    data = request.get_json() or {}
    items = data.get('items', [])
    tax_type = data.get('tax_type', 'IGST')
    delivery_charge = safe_float(data.get('delivery_charge', 0), 0.0)
    tax_rate_igst = safe_float(data.get('tax_rate_igst', 5.0), 5.0)
    tax_rate_cgst = safe_float(data.get('tax_rate_cgst', 2.5), 2.5)
    tax_rate_sgst = safe_float(data.get('tax_rate_sgst', 2.5), 2.5)

    totals = calculate_invoice_totals(
        items,
        tax_type,
        delivery_charge=delivery_charge,
        tax_rate_igst=tax_rate_igst,
        tax_rate_cgst=tax_rate_cgst,
        tax_rate_sgst=tax_rate_sgst,
    )
    return jsonify(totals)


@app.route('/calculate_preview', methods=['POST'])
def calculate_preview_route():
    """Backward-compatible alias used by older templates."""
    return calculate_preview()


@app.route('/cleanup_profiles', methods=['POST'])
def cleanup_profiles():
    """Remove duplicate and invalid buyer profiles."""
    db = get_sheets_db_or_none()
    if db is None:
        flash("Set SPREADSHEET_ID and Google credentials before cleaning up profiles.", "error")
        return redirect(url_for('list_profiles'))

    buyer_profiles = db.get_all_buyers()
    original_count = len(buyer_profiles)
    valid_profiles = [p for p in buyer_profiles if p.get('profile_id') and p.get('buyer_name')]

    deduped = []
    seen_ids = set()
    for profile in valid_profiles:
        profile_id = profile.get('profile_id')
        if profile_id in seen_ids:
            continue
        seen_ids.add(profile_id)
        deduped.append(profile)

    deduped.sort(key=lambda p: p.get('buyer_name', '').lower())

    if len(deduped) < original_count:
        flash(f"Cleanup complete. Removed {original_count - len(deduped)} duplicate/invalid profiles.", "success")
    else:
        flash("No duplicate or invalid profiles found.", "success")

    return redirect(url_for('list_profiles'))


def normalize_date_yyyy_mm_dd(date_str: Any) -> str:
    """Normalize any date format to standard HTML5 input value YYYY-MM-DD."""
    if not date_str:
        return ''
    date_str = str(date_str).strip()
    if not date_str:
        return ''
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return date_str
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y', '%Y/%m/%d'):
        try:
            return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    return date_str


def _enrich_invoice_for_frontend(invoice: Dict[str, Any]) -> Dict[str, Any]:
    """Ensure complete metadata parity across all loading paths."""
    safe_name = invoice.get('filename') or _filename_from_storage_url(invoice.get('file_url', ''), '.xlsx')
    if not safe_name and invoice.get('invoice_number'):
        inv_clean = _safe_filename(str(invoice.get('invoice_number')).replace('/', '-'), '')
        try:
            storage = get_cloud_storage()
            blobs = list(storage.client.list_blobs(storage.bucket_name, prefix=f"{storage.INVOICES_FOLDER}Invoice_{inv_clean}"))
            if blobs:
                safe_name = blobs[0].name.replace(storage.INVOICES_FOLDER, '')
        except Exception:
            pass

    if safe_name:
        invoice['filename'] = safe_name
        invoice['pdf_filename'] = safe_name.replace('.xlsx', '.pdf')
        try:
            storage = get_cloud_storage()
            file_bytes = storage.download_invoice_xlsx(safe_name)
            if file_bytes:
                file_data = _extract_invoice_data_from_xlsx_bytes(file_bytes, safe_name)
                if file_data.get('delivery_charge') is not None:
                    invoice['delivery_charge'] = safe_float(file_data.get('delivery_charge'), 0.0)
                if file_data.get('invoice_date'):
                    invoice['invoice_date'] = file_data['invoice_date']
                if file_data.get('ewaybill_number'):
                    invoice['ewaybill_number'] = file_data['ewaybill_number']
                if file_data.get('ewaybill_date'):
                    invoice['ewaybill_date'] = file_data['ewaybill_date']
                if file_data.get('items'):
                    invoice['items'] = file_data['items']
                if file_data.get('transport_mode'):
                    invoice['transport_mode'] = file_data['transport_mode']
                if file_data.get('ship_from_details'):
                    invoice['ship_from_details'] = file_data['ship_from_details']
                if file_data.get('ship_to_details'):
                    invoice['ship_to_details'] = file_data['ship_to_details']
                if file_data.get('tax_type'):
                    invoice['tax_type'] = file_data['tax_type']
                    invoice['display_tax_type'] = file_data['tax_type']
        except Exception as exc:
            app.logger.warning("Could not enrich invoice from XLSX %s: %s", safe_name, exc)

    # Standardize dates for HTML5 date input
    invoice['invoice_date'] = normalize_date_yyyy_mm_dd(invoice.get('invoice_date', ''))
    invoice['ewaybill_date'] = normalize_date_yyyy_mm_dd(invoice.get('ewaybill_date', ''))
    invoice['delivery_charge'] = safe_float(invoice.get('delivery_charge', 0.0), 0.0)
    invoice['transport_mode'] = extract_transport_core(invoice.get('transport_mode', ''))

    # Parse bags from items descriptions if needed
    for item in invoice.get('items', []):
        desc = item.get('description', '')
        bag_match = re.search(r'\(\s*(\d+(?:\.\d+)?)\s*Bags?\s*\)', desc, re.IGNORECASE)
        if bag_match and not item.get('bags'):
            item['bags'] = bag_match.group(1)
            item['description'] = re.sub(r'\s*\(\s*\d+(?:\.\d+)?\s*Bags?\s*\)', '', desc, flags=re.IGNORECASE).strip()

    # Buyer profile match
    try:
        db = get_sheets_db_or_none()
        if db:
            buyers = db.get_all_buyers()
            matched = _match_buyer_profile(invoice, buyers)
            if matched:
                invoice['buyer_profile_id'] = matched.get('profile_id', '')
                invoice['buyer_name'] = matched.get('buyer_name') or invoice.get('buyer_name', '')
                invoice['buyer_gstin'] = matched.get('gstin') or invoice.get('buyer_gstin', '')
                invoice['buyer_details'] = matched.get('buyer_details', invoice.get('buyer_details', []))
    except Exception as exc:
        app.logger.warning("Could not match buyer profile: %s", exc)

    return invoice


@app.route('/api/invoices')
def api_get_invoices():
    """API endpoint to get latest invoice records in real-time, deduplicated by invoice number."""
    db = get_sheets_db_or_none()
    try:
        sheet_invoices = db.get_all_invoices(limit=2000) if db else []
    except Exception as exc:
        app.logger.warning("Could not fetch invoices: %s", exc)
        sheet_invoices = []

    recent_invoices = _build_invoice_rows(sheet_invoices)
    try:
        storage_rows = get_cloud_storage().list_invoices(limit=2000)
        recent_invoices = _merge_with_storage_rows(recent_invoices, storage_rows)
    except Exception as exc:
        app.logger.warning("Could not list storage invoices: %s", exc)

    # Deduplicate invoices by unique file or distinct bill signature so multiple bills sharing
    # the same sequential number (e.g. 61/2026-27) are preserved individually in the past records list.
    deduped = {}
    for inv in recent_invoices:
        inv_num = str(inv.get('invoice_number', '')).strip()
        if not inv_num:
            continue
        fn = str(inv.get('filename') or '').strip().lower()
        if fn:
            unique_key = fn
        else:
            b_name = str(inv.get('buyer_name') or '').strip().lower()
            i_date = str(inv.get('invoice_date') or '').strip()
            i_amt = str(inv.get('total_amount') or '')
            unique_key = f"{inv_num.lower()}_{b_name}_{i_date}_{i_amt}"

        if unique_key not in deduped:
            deduped[unique_key] = inv
        else:
            prev = deduped[unique_key]
            if (len(inv.get('items') or []) > len(prev.get('items') or [])) or (inv.get('file_url') and not prev.get('file_url')):
                deduped[unique_key] = inv

    deduped_list = list(deduped.values())
    deduped_list.sort(key=_invoice_sort_key, reverse=True)
    return jsonify({'success': True, 'invoices': deduped_list})


@app.route('/api/invoice/<path:invoice_number>')
def api_get_invoice(invoice_number):
    """API endpoint to get invoice data for loading."""
    target_clean = (invoice_number or '').strip()
    if target_clean.lower().endswith('.xlsx'):
        return api_get_invoice_by_file(target_clean)

    db = get_sheets_db_or_none()
    invoice = None
    if db is not None:
        try:
            invoice = db.get_invoice(target_clean)
        except Exception as exc:
            app.logger.warning("Could not fetch invoice from Sheets: %s", exc)

    if not invoice:
        # Fallback search by filename or bill number in GCS
        pattern = _safe_filename(invoice_number.replace('/', '-'), '')
        num_m = re.search(r'(\d+)', invoice_number)
        num_val = int(num_m.group(1)) if num_m else None
        storage = get_cloud_storage()
        try:
            blobs = list(storage.client.list_blobs(storage.bucket_name, prefix=f"{storage.INVOICES_FOLDER}Invoice_"))
            for b in blobs:
                name = b.name.replace(storage.INVOICES_FOLDER, '')
                if pattern and pattern in name:
                    file_bytes = storage.download_invoice_xlsx(name)
                    if file_bytes:
                        invoice = _extract_invoice_data_from_xlsx_bytes(file_bytes, name)
                        break
                elif num_val is not None:
                    m = re.search(r'Invoice_0*(\d+)', name)
                    if m and int(m.group(1)) == num_val:
                        file_bytes = storage.download_invoice_xlsx(name)
                        if file_bytes:
                            invoice = _extract_invoice_data_from_xlsx_bytes(file_bytes, name)
                            break
        except Exception as exc:
            app.logger.warning("Error during GCS fallback for %s: %s", invoice_number, exc)

    if not invoice:
        return jsonify({'error': 'Invoice not found', 'success': False}), 404

    invoice = _enrich_invoice_for_frontend(invoice)
    resp = dict(invoice)
    resp['invoice'] = invoice
    resp['success'] = True
    return jsonify(resp)


@app.route('/api/invoice-file/<path:filename>')
def api_get_invoice_by_file(filename):
    """Load invoice data directly from a stored XLSX file."""
    safe_name = _safe_filename(filename, '.xlsx')
    file_bytes = get_cloud_storage().download_invoice_xlsx(safe_name)
    if not file_bytes:
        return jsonify({'error': 'Invoice file not found'}), 404

    try:
        invoice = _extract_invoice_data_from_xlsx_bytes(file_bytes, safe_name)
    except Exception as exc:
        app.logger.exception('Failed to parse XLSX %s: %s', safe_name, exc)
        return jsonify({'error': 'Failed to parse invoice file'}), 500

    invoice['filename'] = safe_name
    invoice['pdf_filename'] = safe_name.replace('.xlsx', '.pdf')
    invoice = _enrich_invoice_for_frontend(invoice)
    return jsonify(invoice)


@app.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    """Retrieve or update dynamic application settings."""
    storage = get_cloud_storage()
    if request.method == 'POST':
        updates = request.get_json(silent=True) or {}
        # Retain server key if empty/whitespace string was submitted
        if 'gemini_backend_api_key' in updates:
            provided_key = str(updates['gemini_backend_api_key']).strip()
            if not provided_key:
                current = get_settings(storage)
                if current.get('gemini_backend_api_key'):
                    updates['gemini_backend_api_key'] = current['gemini_backend_api_key']
        new_settings = update_settings(updates, storage)
        return jsonify({'success': True, 'settings': sanitize_settings_for_client(new_settings)})
    return jsonify({'success': True, 'settings': sanitize_settings_for_client(get_settings(storage))})


def compress_image_b64(raw_b64: str, max_dimension: int = 1024, quality: int = 72) -> Tuple[str, str]:
    """Ensure base64 image data does not overload Gemini API while preserving crisp legibility for OCR."""
    try:
        from PIL import Image
        img_bytes = base64.b64decode(raw_b64)
        with Image.open(io.BytesIO(img_bytes)) as im:
            if im.mode in ('RGBA', 'P', 'LA'):
                im = im.convert('RGB')
            w, h = im.size
            if w > max_dimension or h > max_dimension:
                if w > h:
                    h = max(1, int(h * max_dimension / w))
                    w = max_dimension
                else:
                    w = max(1, int(w * max_dimension / h))
                    h = max_dimension
                im = im.resize((w, h), Image.Resampling.LANCZOS)
            out_buf = io.BytesIO()
            im.save(out_buf, format='JPEG', quality=quality, optimize=True)
            new_b64 = base64.b64encode(out_buf.getvalue()).decode('utf-8')
            return new_b64, 'image/jpeg'
    except Exception as e:
        app.logger.warning("Could not server-side compress image: %s", e)
        return raw_b64, 'image/jpeg'


@app.route('/api/ai/parse-bill', methods=['POST'])
def api_ai_parse_bill():
    """
    Extract invoice information from multiple uploaded pictures (rough slips, chits, visiting cards)
    and/or user notes using Google Gemini Flash.
    Designed for elderly-friendly usage with free-tier rate limit efficiency.
    """
    storage = get_cloud_storage()
    current_settings = get_settings(storage)

    data = request.get_json(silent=True) or {}
    images = data.get('images', [])  # List of {data: base64_str, mime_type: 'image/jpeg'}
    session_images = data.get('session_images', [])
    if not images and session_images:
        images = session_images
    prompt_text = data.get('prompt', '').strip()
    create_buyer_if_new = data.get('create_new_buyer', True)

    api_key = (
        request.headers.get('X-Gemini-Key') or 
        data.get('api_key') or 
        os.environ.get('GEMINI_API_KEY', '') or
        current_settings.get('gemini_backend_api_key', '') or
        DEFAULT_SETTINGS.get('gemini_backend_api_key', '')
    ).strip()

    if not api_key:
        return jsonify({
            'success': False,
            'error': 'Gemini API key is required. Please configure it in Settings or get a free key at https://aistudio.google.com/apikey.'
        }), 400

    if not images and not prompt_text:
        return jsonify({'success': False, 'error': 'Please provide at least one image or typed note.'}), 400

    # Retrieve existing buyers, transports, dispatch addresses, and invoice history for grounding
    known_buyers = []
    known_transports = []
    known_dispatch_addresses = []
    existing_invoices_sample = []
    suggested_inv = suggest_next_invoice_number()

    try:
        db = get_sheets_db_or_none()
        if db:
            all_b = db.get_all_buyers()
            known_buyers = [
                {
                    'profile_id': b.get('profile_id'),
                    'buyer_name': b.get('buyer_name'),
                    'gstin': b.get('gstin', ''),
                    'state': b.get('state', ''),
                    'state_code': b.get('state_code', ''),
                    'details': b.get('buyer_details', [])
                }
                for b in all_b if b.get('buyer_name')
            ]
            known_transports = db.get_all_transport_modes()
            known_dispatch_addresses = db.get_all_dispatch_addresses()
            past_invs = db.get_all_invoices(limit=200)
            existing_invoices_sample = [inv.get('invoice_number', '').strip() for inv in past_invs if inv.get('invoice_number')]
    except Exception as e:
        app.logger.warning("Could not fetch DB records for AI context: %s", e)

    if not known_dispatch_addresses:
        known_dispatch_addresses = list(getattr(GoogleSheetsDB, 'DEFAULT_DISPATCH_ADDRESSES', []))

    # Also check Cloud Storage for invoice numbers
    if storage:
        try:
            st_invs = storage.list_invoices(limit=200)
            for si in st_invs:
                num = si.get('invoice_number', '').strip()
                if num and num not in existing_invoices_sample:
                    existing_invoices_sample.append(num)
        except Exception:
            pass

    history = data.get('history', [])  # Multi-turn conversation history
    current_draft = data.get('current_draft') or {}
    current_invoices = data.get('current_invoices') or []
    current_form = data.get('current_form') or {}

    processed_images = []
    for img in images:
        raw_b64 = img.get('data', '')
        if ',' in raw_b64:
            raw_b64 = raw_b64.split(',', 1)[1]
        mtype = img.get('mime_type', 'image/jpeg')
        if raw_b64:
            comp_b64, comp_mtype = compress_image_b64(raw_b64, max_dimension=1024, quality=72)
            processed_images.append({'mime_type': comp_mtype, 'data': comp_b64})

    processed_session_images = []
    for simg in session_images:
        raw_b64 = simg.get('data', '')
        if ',' in raw_b64:
            raw_b64 = raw_b64.split(',', 1)[1]
        mtype = simg.get('mime_type', 'image/jpeg')
        if raw_b64:
            comp_b64, comp_mtype = compress_image_b64(raw_b64, max_dimension=1024, quality=72)
            processed_session_images.append({'mime_type': comp_mtype, 'data': comp_b64})

    company_name = current_settings.get('company_name', 'Shakambhari Enterprises')
    buyers_summary = [
        f"{b['buyer_name']} | GSTIN: {b.get('gstin', 'N/A')} | State: {b.get('state', '')} ({b.get('state_code', '')}) | ID: {b['profile_id']}"
        for b in known_buyers
    ]

    instruction = (
        f"You are the senior billing assistant and dispatch agent for {company_name}, an aluminium utensils manufacturing and trading business in Liluah/Howrah, West Bengal.\n"
        "You operate agentically: you read rough paper slips/chits, extract invoices, update or edit buyers, modify active bill fields upon user request, and maintain conversational context across multiple turns.\n\n"
        "=== DATABASE GROUNDING (YOUR KNOWLEDGE BASE) ===\n"
        f"1. KNOWN BUYERS IN DATABASE ({len(known_buyers)} profiles):\n"
        f"{json.dumps(buyers_summary, ensure_ascii=False, indent=1)}\n\n"
        f"2. KNOWN TRANSPORTS:\n"
        f"{json.dumps(known_transports, ensure_ascii=False)}\n\n"
        f"3. KNOWN DISPATCH FROM ADDRESSES:\n"
        f"{json.dumps(known_dispatch_addresses, ensure_ascii=False)}\n\n"
        f"4. RECENT INVOICE NUMBERS ALREADY USED IN DATABASE:\n"
        f"{json.dumps(existing_invoices_sample[:80], ensure_ascii=False)}\n\n"
        f"5. SUGGESTED NEXT SEQUENTIAL INVOICE NUMBER: {suggested_inv}\n\n"
        "=== AGENTIC TOOLS AVAILABLE TO YOU ===\n"
        "1. `save_or_update_buyer`:\n"
        "   Use this tool when the user asks to edit an existing buyer's details (address, GSTIN, Aadhaar, state, name) OR save a new buyer profile into the database.\n"
        "   Args: {\"profile_id\": \"<existing ID or null>\", \"buyer_name\": \"...\", \"buyer_details\": [\"...\"], \"gstin\": \"...\", \"aadhaar\": \"...\", \"state\": \"...\", \"state_code\": \"...\", \"default_tax_type\": \"IGST\" or \"CGST_SGST\"}\n\n"
        "2. `edit_invoice_bill`:\n"
        "   Use this tool when the user asks to modify an active bill or field (e.g. rate, bags, weight, transport, date, buyer, delivery charge, item description).\n"
        "   Args: {\"bill_index\": 0, \"fields_to_update\": { ... }}\n\n"
        "=== AGENT INSTRUCTIONS & RULES ===\n"
        "1. Multi-bill slips: If a single paper chit contains multiple bills for different parties, extract each bill into the 'invoices' array. Sequential invoice numbers should start with suggested number and increment for subsequent bills.\n"
        "2. Multi-turn conversation: Maintain conversational state across turns. When the user asks to modify, correct, or adjust any field (e.g. rate, bags, transport, buyer), apply the requested change while preserving all other extracted details.\n"
        "3. Buyer profile editing: If the user asks to update a buyer's address, GSTIN, Aadhaar, or state, invoke `save_or_update_buyer` and update the active bill.\n"
        "4. Tolerant buyer matching: Match known buyers flexibly. If a party is new or uses Aadhaar, extract person name as buyer_name, omit father/care-of names, store the 12-digit number in 'aadhaar', leave 'gstin' empty '', and format the last line of buyer_details as 'AADHAAR - XXXX XXXX XXXX'.\n"
        "5. Packaging units & items: Default unit to 'Bags'. Default HSN to '76151030'. Ensure quantities, rates, and amounts are numeric.\n"
        "6. Strict Output: Always respond with STRICT JSON ONLY matching the schema below (NO markdown code fences, NO extraneous commentary):\n\n"
        "{\n"
        '  "conversational_message": "Warm, plain Hinglish/English explanation of what was extracted or modified...",\n'
        '  "tools_called": [\n'
        '    {\n'
        '      "name": "save_or_update_buyer" or "edit_invoice_bill",\n'
        '      "args": { ... }\n'
        '    }\n'
        '  ],\n'
        '  "has_clarifications": false,\n'
        '  "clarifications": [],\n'
        '  "is_duplicate_invoice": false,\n'
        '  "suggested_next_invoice_number": "' + suggested_inv + '",\n'
        '  "invoices": [\n'
        '    {\n'
        '      "invoice_number": "' + suggested_inv + '",\n'
        '      "invoice_date": "' + get_ist_today_str() + '",\n'
        '      "ewaybill_number": "",\n'
        '      "ewaybill_date": "",\n'
        '      "buyer_name": "...",\n'
        '      "matched_profile_id": "..." or null,\n'
        '      "is_new_buyer": false,\n'
        '      "buyer_details": ["Line 1", "Line 2"],\n'
        '      "gstin": "...",\n'
        '      "aadhaar": "...",\n'
        '      "state": "...",\n'
        '      "state_code": "...",\n'
        '      "transport_mode": "...",\n'
        '      "dispatch_from": "...",\n'
        '      "delivery_charge": 0.0,\n'
        '      "tax_type": "IGST" or "CGST_SGST",\n'
        '      "items": [\n'
        '        {\n'
        '          "description": "Aluminium Utensils",\n'
        '          "bags": "2",\n'
        '          "unit": "Bags",\n'
        '          "quantity": 84.0,\n'
        '          "rate": 380.00,\n'
        '          "hsn": "76151030"\n'
        '        }\n'
        '      ]\n'
        '    }\n'
        '  ],\n'
        '  "data": { ...first bill object... }\n'
        "}"
    )

    contents = []

    # If multi-turn chat history is provided, reconstruct proper alternating conversational turns
    if history and isinstance(history, list):
        for h_turn in history:
            role = h_turn.get('role')
            if role == 'user':
                t_parts = []
                # Ground the first user turn with session images if available
                if len(contents) == 0 and processed_session_images:
                    for pimg in processed_session_images:
                        t_parts.append({
                            'inline_data': {
                                'mime_type': pimg['mime_type'],
                                'data': pimg['data']
                            }
                        })
                t_text = h_turn.get('text', '').strip()
                if t_text:
                    t_parts.append({'text': t_text})
                elif not t_parts:
                    t_parts.append({'text': 'Slip image uploaded.'})
                contents.append({'role': 'user', 'parts': t_parts})
            elif role == 'model':
                m_text = h_turn.get('text', '').strip()
                if not m_text and h_turn.get('raw_response'):
                    m_text = str(h_turn.get('raw_response')).strip()
                if not m_text:
                    m_text = json.dumps({'conversational_message': 'Received.'})
                contents.append({'role': 'model', 'parts': [{'text': m_text}]})

    # Now construct the current user turn
    current_parts = []
    # If no history, or if new images were sent in this turn, attach them
    images_for_current = processed_images if processed_images else ([] if history else processed_session_images)
    for pimg in images_for_current:
        current_parts.append({
            'inline_data': {
                'mime_type': pimg['mime_type'],
                'data': pimg['data']
            }
        })

    user_text = prompt_text if prompt_text else ("Please examine the provided slip picture(s) and extract all billing details." if not history else "Please update the bill according to my instructions.")

    # Context grounding for current turn (screen state)
    grounding_notes = []
    if current_form:
        grounding_notes.append(f"CURRENT INVOICE FORM VALUES ON SCREEN:\n{json.dumps(current_form, ensure_ascii=False)}")
    if current_invoices:
        grounding_notes.append(f"CURRENT WORKING BILLS:\n{json.dumps(current_invoices, ensure_ascii=False)}")
    elif current_draft:
        grounding_notes.append(f"CURRENT WORKING DRAFT:\n{json.dumps(current_draft, ensure_ascii=False)}")

    if grounding_notes:
        full_turn_text = "\n\n".join(grounding_notes) + f"\n\nUSER REQUEST:\n{user_text}"
    else:
        full_turn_text = user_text

    current_parts.append({'text': full_turn_text})
    contents.append({'role': 'user', 'parts': current_parts})

    gemini_payload = {
        "systemInstruction": {
            "parts": [{"text": instruction}]
        },
        "contents": contents,
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json"
        }
    }

    # Build resilient model hierarchy using active, verified Gemini 3.x models
    # Strictly exclude decommissioned models (1.5, 2.0, 2.5) and high-load aliases (gemini-flash-latest)
    disallowed_models = {
        "gemini-flash-latest", "gemini-2.5-flash", "gemini-2.5-flash-lite",
        "gemini-2.5-pro", "gemini-2.0-flash", "gemini-2.0-flash-001",
        "gemini-2.0-flash-lite", "gemini-1.5-flash", "gemini-1.5-flash-8b",
        "gemini-1.5-pro", "gemini-flash", "gemini-pro"
    }
    standard_models_pool = [
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3-flash-preview",
        "gemini-3.1-flash-lite",
        "gemini-3.1-flash-lite-preview",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite"
    ]

    primary_m = current_settings.get('gemini_primary_model', 'gemini-3.8-flash')
    if primary_m in disallowed_models or not primary_m:
        primary_m = 'gemini-3.8-flash'

    raw_backups = current_settings.get('gemini_backup_models') or []
    configured_backups = [m for m in raw_backups if m not in disallowed_models and m != primary_m]

    # Combine into ordered, deduplicated fallback queue with guaranteed full pool coverage
    models = [primary_m]
    for m in configured_backups:
        if m not in models:
            models.append(m)
    for m in standard_models_pool:
        if m not in models:
            models.append(m)

    last_error = ""

    for model in models:
        text_out = None
        api_url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        try:
            req = urllib.request.Request(
                api_url,
                data=json.dumps(gemini_payload).encode('utf-8'),
                headers={'Content-Type': 'application/json'},
                method='POST'
            )
            with urllib.request.urlopen(req, timeout=18) as resp:
                res_body = resp.read().decode('utf-8')
                res_data = json.loads(res_body)
                candidates = res_data.get('candidates', [])
                if candidates:
                    content = candidates[0].get('content', {})
                    resp_parts = content.get('parts', [])
                    if resp_parts:
                        text_out = resp_parts[0].get('text', '').strip()
                if not text_out:
                    last_error = f"Model {model} generateContent returned empty output."
        except urllib.error.HTTPError as he:
            err_msg = he.read().decode('utf-8', errors='ignore')
            app.logger.warning("Gemini model %s HTTP %s: %s", model, he.code, err_msg)
            last_error = f"{model} HTTP {he.code}: {err_msg[:120]}"
            time.sleep(0.3)
        except Exception as exc:
            app.logger.warning("Gemini model %s error: %s", model, exc)
            last_error = f"{model} error: {exc}"

        if not text_out:
            continue

        # Clean code fences if present
        clean_json_str = re.sub(r'^```(?:json)?\s*', '', text_out, flags=re.MULTILINE)
        clean_json_str = re.sub(r'\s*```$', '', clean_json_str, flags=re.MULTILINE).strip()
        try:
            parsed_response = json.loads(clean_json_str)
        except Exception:
            m_json = re.search(r'(\{[\s\S]*\})', clean_json_str)
            if m_json:
                repaired = re.sub(r',\s*([}\]])', r'\1', m_json.group(1))
                try:
                    parsed_response = json.loads(repaired)
                except Exception:
                    last_error = f"Model {model} produced unparseable JSON: {clean_json_str[:120]}"
                    continue
            else:
                last_error = f"Model {model} produced unparseable JSON: {clean_json_str[:120]}"
                continue

        raw_invoices = parsed_response.get('invoices')
        if raw_invoices is None:
            single_val = parsed_response.get('data')
            if isinstance(single_val, list):
                raw_invoices = single_val
            elif isinstance(single_val, dict) and (single_val.get('buyer_name') or single_val.get('items')):
                raw_invoices = [single_val]
            else:
                raw_invoices = []
        elif not isinstance(raw_invoices, list):
            raw_invoices = []

        today_ist = get_ist_today_str()
        parsed_invoices = []

        for idx, inv in enumerate(raw_invoices):
            if not isinstance(inv, dict):
                continue

            b_name = (inv.get('buyer_name') or '').strip()
            raw_items = inv.get('items') or []
            has_real_buyer = bool(b_name and b_name.lower() not in ('unnamed party', 'unknown', 'none', 'n/a', 'null', ''))
            has_real_items = any(
                (float(it.get('quantity') or 0) > 0 or float(it.get('rate') or 0) > 0 or (it.get('description') and it.get('description').strip()))
                for it in raw_items if isinstance(it, dict)
            ) if isinstance(raw_items, list) else False

            # Skip empty dummy bill if neither a real party nor items exist
            if not has_real_buyer and not has_real_items:
                continue

            inv_data = dict(inv)
            if not inv_data.get('invoice_date'):
                inv_data['invoice_date'] = today_ist

            # Assign sequential invoice number if missing
            if not inv_data.get('invoice_number') and suggested_inv:
                inv_data['invoice_number'] = format_sequential_invoice_number(suggested_inv, idx)

            # Normalize buyer profile info on invoice
            if inv_data.get('buyer_name') or inv_data.get('buyer_details'):
                temp_p = {
                    'buyer_name': inv_data.get('buyer_name', ''),
                    'buyer_details': inv_data.get('buyer_details') or [inv_data.get('buyer_name', '')],
                    'gstin': inv_data.get('gstin', ''),
                    'aadhaar': inv_data.get('aadhaar') or inv_data.get('aadhar', ''),
                    'state': inv_data.get('state', ''),
                    'state_code': inv_data.get('state_code', ''),
                    'default_tax_type': inv_data.get('tax_type', 'IGST')
                }
                norm_p = normalize_buyer_profile(temp_p)
                inv_data['buyer_details'] = norm_p['buyer_details']
                inv_data['gstin'] = norm_p['gstin']
                inv_data['aadhaar'] = norm_p.get('aadhaar', '')
                inv_data['state'] = norm_p['state']
                inv_data['state_code'] = norm_p['state_code']
                inv_data['tax_type'] = norm_p['default_tax_type']

            # Auto-persist new buyer if requested
            if inv_data.get('is_new_buyer') and inv_data.get('buyer_name') and create_buyer_if_new:
                b_name = inv_data.get('buyer_name', '').strip()
                matched = next((kb for kb in known_buyers if kb.get('buyer_name', '').strip().lower() == b_name.lower()), None)
                if matched:
                    inv_data['is_new_buyer'] = False
                    inv_data['matched_profile_id'] = matched.get('profile_id')
                else:
                    new_profile_id = f"buyer_{int(time.time())}_{uuid.uuid4().hex[:4]}"
                    new_profile_raw = {
                        'profile_id': new_profile_id,
                        'buyer_name': b_name,
                        'buyer_details': inv_data.get('buyer_details') or [b_name],
                        'gstin': inv_data.get('gstin', '').strip().upper(),
                        'aadhaar': inv_data.get('aadhaar', ''),
                        'state': inv_data.get('state', ''),
                        'state_code': inv_data.get('state_code', ''),
                        'default_tax_type': inv_data.get('tax_type', 'IGST')
                    }
                    new_profile = normalize_buyer_profile(new_profile_raw)
                    try:
                        db = get_sheets_db_or_none()
                        if db:
                            db.save_buyer(new_profile)
                            known_buyers.append(new_profile)
                    except Exception as err:
                        app.logger.warning("Could not auto-save new buyer profile: %s", err)
                    inv_data['matched_profile_id'] = new_profile_id
                    inv_data['created_profile'] = new_profile

            # Normalize items: guarantee default description, Bags unit, and numeric types
            norm_items = []
            raw_items = inv_data.get('items', [])
            if isinstance(raw_items, list):
                for it in raw_items:
                    if isinstance(it, dict):
                        norm_items.append({
                            'description': it.get('description') or 'Aluminium Utensils',
                            'bags': str(it.get('bags') or ''),
                            'unit': it.get('unit') or 'Bags',
                            'quantity': float(it.get('quantity') or 0),
                            'rate': float(it.get('rate') or 0),
                            'hsn': it.get('hsn') or '76151030'
                        })
            inv_data['items'] = norm_items
            parsed_invoices.append(inv_data)

        # Execute agentic tools if returned by the model
        saved_buyers_list = []
        tools_called = parsed_response.get('tools_called', [])
        if isinstance(tools_called, list):
            for tc in tools_called:
                if isinstance(tc, dict) and tc.get('name') == 'save_or_update_buyer':
                    args = tc.get('args') or {}
                    b_name = (args.get('buyer_name') or '').strip()
                    if b_name:
                        p_id = (args.get('profile_id') or '').strip()
                        if not p_id:
                            m_kb = next((kb for kb in known_buyers if kb.get('buyer_name', '').strip().lower() == b_name.lower()), None)
                            p_id = m_kb.get('profile_id') if m_kb else f"buyer_{int(time.time())}_{uuid.uuid4().hex[:4]}"
                        
                        raw_bp = {
                            'profile_id': p_id,
                            'buyer_name': b_name,
                            'buyer_details': args.get('buyer_details') or [b_name],
                            'gstin': (args.get('gstin') or '').strip().upper(),
                            'aadhaar': (args.get('aadhaar') or '').strip(),
                            'state': (args.get('state') or '').strip(),
                            'state_code': (args.get('state_code') or '').strip(),
                            'default_tax_type': args.get('default_tax_type') or 'IGST'
                        }
                        norm_bp = normalize_buyer_profile(raw_bp)
                        try:
                            db = get_sheets_db_or_none()
                            if db:
                                db.save_buyer(norm_bp)
                                saved_buyers_list.append(norm_bp)
                                known_buyers = [kb for kb in known_buyers if kb.get('profile_id') != p_id]
                                known_buyers.append(norm_bp)
                        except Exception as b_err:
                            app.logger.warning("Agent tool save_or_update_buyer error: %s", b_err)

        # Retain previous working bills if the user requested conversational tweaks and no new bills were returned
        if not parsed_invoices and (current_invoices or current_draft) and not parsed_response.get('invoices'):
            parsed_invoices = current_invoices if current_invoices else ([current_draft] if current_draft else [])

        primary_invoice = parsed_invoices[0] if parsed_invoices else None

        return jsonify({
            'success': True,
            'model': model,
            'conversational_message': parsed_response.get('conversational_message', 'Extracted details from slip.'),
            'tools_called': tools_called,
            'saved_buyers': saved_buyers_list,
            'has_clarifications': parsed_response.get('has_clarifications', False),
            'clarifications': parsed_response.get('clarifications', []),
            'is_duplicate_invoice': parsed_response.get('is_duplicate_invoice', False),
            'duplicate_warning': parsed_response.get('duplicate_warning'),
            'suggested_next_invoice_number': parsed_response.get('suggested_next_invoice_number', suggested_inv),
            'invoices': parsed_invoices,
            'data': primary_invoice,
            'raw_response': text_out
        })

    return jsonify({
        'success': False,
        'error': f'All active Gemini models ({len(models)} models attempted) were busy or temporarily unavailable. Please retry in a few moments. Last error: {last_error}'
    }), 503


@app.route('/api/buyer/save', methods=['POST'])
@app.route('/api/ai/create-buyer', methods=['POST'])
def api_save_buyer():
    """Create or update a buyer profile in Google Sheets and memory directory."""
    data = request.get_json(silent=True) or {}
    buyer_name = (data.get('buyer_name') or '').strip()
    if not buyer_name:
        return jsonify({'success': False, 'error': 'Buyer name is required.'}), 400

    profile_id = (data.get('profile_id') or '').strip()
    if not profile_id:
        profile_id = f"buyer_{int(time.time())}_{uuid.uuid4().hex[:4]}"

    details = data.get('buyer_details')
    if isinstance(details, str):
        details = [line.strip() for line in details.splitlines() if line.strip()]
    elif not isinstance(details, list):
        details = [buyer_name]

    gstin = (data.get('gstin') or '').strip().upper()
    state = (data.get('state') or '').strip()
    state_code = (data.get('state_code') or '').strip()
    if not state_code and len(gstin) >= 2 and gstin[:2].isdigit():
        state_code = gstin[:2]

    tax_type = data.get('default_tax_type') or data.get('tax_type') or ('CGST_SGST' if state_code == '19' else 'IGST')

    raw_profile = {
        'profile_id': profile_id,
        'buyer_name': buyer_name,
        'buyer_details': details,
        'gstin': gstin,
        'aadhaar': (data.get('aadhaar') or data.get('aadhar') or '').strip(),
        'state': state,
        'state_code': state_code,
        'default_tax_type': tax_type
    }
    normalized = normalize_buyer_profile(raw_profile)

    try:
        db = get_sheets_db_or_none()
        if db:
            db.save_buyer(normalized)
    except Exception as exc:
        app.logger.warning("Failed to save buyer to Sheets: %s", exc)

    # Persist locally to buyer_profiles.json files if available
    for json_path in ['buyer_profiles.json', 'cloud/buyer_profiles.json']:
        try:
            if os.path.exists(json_path):
                with open(json_path, 'r', encoding='utf-8') as f:
                    local_buyers = json.load(f)
                b_idx = next((i for i, b in enumerate(local_buyers) if b.get('profile_id') == normalized['profile_id']), -1)
                if b_idx >= 0:
                    local_buyers[b_idx] = normalized
                else:
                    local_buyers.append(normalized)
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(local_buyers, f, indent=2, ensure_ascii=False)
        except Exception as file_exc:
            app.logger.warning("Failed to save buyer to %s: %s", json_path, file_exc)

    return jsonify({
        'success': True,
        'buyer': normalized,
        'message': f"Saved '{buyer_name}' to directory."
    })

api_ai_create_buyer = api_save_buyer


@app.route('/api/buyer/<profile_id>', methods=['GET'])
def api_get_buyer(profile_id):
    """Retrieve a single buyer profile."""
    try:
        db = get_sheets_db_or_none()
        if db:
            b = db.get_buyer(profile_id)
            if b:
                return jsonify({'success': True, 'buyer': normalize_buyer_profile(b)})
    except Exception as exc:
        app.logger.warning("Failed to load buyer %s: %s", profile_id, exc)
    return jsonify({'success': False, 'error': 'Buyer not found'}), 404


@app.route('/api/generate_batch', methods=['POST'])
def api_generate_batch():
    """
    Batch generate multiple invoices at once from extracted AI chits.
    Saves all records and generates download links for both PDF and Excel.
    """
    data = request.get_json(silent=True) or {}
    raw_invoices = data.get('invoices', [])
    if not raw_invoices:
        return jsonify({'success': False, 'error': 'No invoices provided for batch generation.'}), 400

    db = get_sheets_db_or_none()
    storage = get_cloud_storage()
    if not db or not storage:
        return jsonify({'success': False, 'error': 'Database or storage service unavailable.'}), 500

    base_suggested_inv = suggest_next_invoice_number()
    results = []
    errors = []

    for idx, inv in enumerate(raw_invoices):
        try:
            buyer = None
            if inv.get('matched_profile_id'):
                buyer = db.get_buyer(inv['matched_profile_id'])
            if not buyer and inv.get('buyer_name'):
                all_b = db.get_all_buyers()
                b_name_lower = inv['buyer_name'].strip().lower()
                buyer = next((b for b in all_b if b.get('buyer_name', '').strip().lower() == b_name_lower), None)
            if not buyer and inv.get('buyer_name'):
                new_id = f"buyer_{int(time.time())}_{uuid.uuid4().hex[:4]}"
                buyer = normalize_buyer_profile({
                    'profile_id': new_id,
                    'buyer_name': inv['buyer_name'].strip(),
                    'buyer_details': inv.get('buyer_details') or [inv['buyer_name'].strip()],
                    'gstin': inv.get('gstin', '').strip().upper(),
                    'state': inv.get('state', ''),
                    'state_code': inv.get('state_code', ''),
                    'default_tax_type': inv.get('tax_type', 'IGST')
                })
                try:
                    db.save_buyer(buyer)
                except Exception:
                    pass

            if not buyer:
                errors.append(f"Bill #{idx+1}: Buyer could not be determined.")
                continue

            inv_number = inv.get('invoice_number') or format_sequential_invoice_number(base_suggested_inv, idx)
            inv_date = inv.get('invoice_date') or get_ist_today_str()
            tax_type = inv.get('tax_type') or buyer.get('default_tax_type', 'IGST')
            transport = normalize_transport_mode(inv.get('transport_mode') or 'By Road')
            delivery = safe_float(inv.get('delivery_charge', 0), 0.0)

            raw_items = inv.get('items', [])
            items = []
            for it in raw_items:
                q = float(it.get('quantity') or 0)
                r = float(it.get('rate') or 0)
                if q > 0 or r > 0:
                    d = (it.get('description') or 'Aluminium Utensils').strip()
                    b = str(it.get('bags') or '').strip()
                    u = (it.get('unit') or 'Bags').strip()
                    d = re.sub(r'\s*\(\s*[\d\.]+\s*(?:[A-Za-z]+|\s*)+\)', '', d, flags=re.IGNORECASE).strip()
                    if b:
                        d += f" ({b} {u})"
                    items.append({
                        'description': d,
                        'bags': b,
                        'unit': u,
                        'hsn': it.get('hsn') or '76151030',
                        'quantity': q,
                        'rate': r,
                        'amount': q * r
                    })

            if not items:
                errors.append(f"Bill #{idx+1} ({buyer['buyer_name']}): No items specified.")
                continue

            totals = calculate_invoice_totals(items, tax_type, delivery_charge=delivery)

            ship_from_text = (inv.get('dispatch_from') or '').strip()
            ship_from_lines = [l.strip() for l in re.split(r'[\r\n]+', ship_from_text) if l.strip()] if ship_from_text else []
            ship_to_lines = buyer.get('buyer_details', [])

            invoice_data = {
                'invoice_number': inv_number,
                'invoice_date': inv_date,
                'invoice_date_display': format_date_for_invoice(inv_date),
                'buyer_name': buyer['buyer_name'],
                'buyer_gstin': buyer.get('gstin', ''),
                'buyer_details': buyer.get('buyer_details', []),
                'ewaybill_number': inv.get('ewaybill_number', ''),
                'ewaybill_date': inv.get('ewaybill_date', ''),
                'ewaybill_date_display': format_date_for_invoice(inv.get('ewaybill_date')) if inv.get('ewaybill_date') else '',
                'ship_from_details': ship_from_lines,
                'ship_to_enabled': True,
                'ship_to_details': ship_to_lines,
                'digitally_signed': True,
                'items': items,
                'transport_mode': transport,
                'delivery_charge': delivery,
                'tax_type': tax_type,
                'display_tax_type': totals['display_tax_type'],
                'tax_rate_igst': totals['igst_rate'],
                'tax_rate_cgst': totals['cgst_rate'],
                'tax_rate_sgst': totals['sgst_rate'],
                **totals
            }

            excel_bytes = generate_invoice_excel(invoice_data)
            safe_buyer = ''.join(c if c.isalnum() else '_' for c in buyer['buyer_name'][:20])
            filename = f"Invoice_{inv_number.replace('/', '-')}_{inv_date}_{safe_buyer}.xlsx"
            xlsx_url = storage.upload_invoice_xlsx(excel_bytes, filename)

            pdf_filename = filename.replace('.xlsx', '.pdf')
            pdf_bytes = generate_pdf_from_excel(excel_bytes, filename)
            pdf_url = ''
            if pdf_bytes:
                pdf_url = storage.upload_invoice_pdf(pdf_bytes, pdf_filename)

            invoice_record = {
                'invoice_number': inv_number,
                'invoice_date': inv_date,
                'buyer_name': buyer['buyer_name'],
                'buyer_gstin': buyer.get('gstin', ''),
                'items': items,
                'subtotal': totals['subtotal'],
                'tax_type': tax_type,
                'display_tax_type': totals['display_tax_type'],
                'tax_rate_igst': totals['igst_rate'],
                'tax_rate_cgst': totals['cgst_rate'],
                'tax_rate_sgst': totals['sgst_rate'],
                'tax_amount': totals['tax_amount'],
                'total_amount': totals['rounded_total'],
                'transport_mode': transport,
                'file_url': xlsx_url,
                'pdf_url': pdf_url,
                'ewaybill_number': inv.get('ewaybill_number', ''),
                'ewaybill_date': inv.get('ewaybill_date', ''),
                'ship_from_details': ship_from_lines,
                'ship_to_enabled': True,
                'ship_to_details': ship_to_lines,
                'digitally_signed': True,
                'delivery_charge': delivery,
            }
            try:
                db.save_invoice(invoice_record)
                if transport:
                    db.add_transport_mode(transport)
                if ship_from_text:
                    db.add_dispatch_address(ship_from_text)
            except Exception as se:
                app.logger.warning("Could not sync to Sheets: %s", se)

            excel_download_url = url_for('download_xlsx', filename=filename)
            pdf_download_url = url_for('download_pdf', filename=pdf_filename) if pdf_url else ''

            results.append({
                'invoice_number': inv_number,
                'buyer_name': buyer['buyer_name'],
                'total_amount': totals['rounded_total'],
                'filename': filename,
                'excel_url': excel_download_url,
                'pdf_url': pdf_download_url,
                'is_pdf': bool(pdf_url)
            })

        except Exception as err:
            app.logger.error("Error generating bill in batch: %s", err)
            errors.append(f"Bill #{idx+1}: {str(err)}")

    if not results:
        return jsonify({'success': False, 'error': '; '.join(errors) or 'Batch generation failed.'}), 400

    return jsonify({
        'success': True,
        'created_count': len(results),
        'results': results,
        'errors': errors
    })


@app.errorhandler(404)
def not_found(_error):
    """Render a friendly 404 page."""
    return render_template('error.html',
                          code=404,
                          title='Page Not Found',
                          message='The page you requested does not exist.'), 404


@app.errorhandler(429)
def too_many_requests(_error):
    """Render a friendly 429 page."""
    return render_template('error.html',
                          code=429,
                          title='Too Many Requests',
                          message='Please slow down and retry after a short pause.'), 429


@app.errorhandler(Exception)
def handle_unexpected_error(error):
    """Fallback error handler with logging for production debugging."""
    if isinstance(error, HTTPException):
        return error

    app.logger.exception('Unhandled error: %s', error)
    return render_template('error.html',
                          code=500,
                          title='Something Went Wrong',
                          message='An unexpected error occurred. Please try again.'), 500


# ===================== HEALTH CHECK =====================

@app.route('/health')
def health_check():
    """Health check endpoint for Cloud Run/App Engine."""
    return jsonify({'status': 'healthy', 'timestamp': get_ist_now().isoformat()})


# ===================== MAIN =====================

if __name__ == '__main__':
    # For local development
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
