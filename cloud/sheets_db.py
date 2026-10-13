from __future__ import annotations
import time
"""
Google Sheets Integration for Shakambhari Invoice Generator
============================================================
This module handles all Google Sheets operations for:
- Buyer Profiles (stored in 'Buyers' sheet)
- Transport Modes (stored in 'Transport' sheet)
- Invoice Records (stored in 'Invoices' sheet)
"""

import os
import re
import json
from datetime import datetime
from typing import List, Dict, Optional
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

try:
    from buyer_normalizer import normalize_buyer_profile
except ImportError:
    try:
        from cloud.buyer_normalizer import normalize_buyer_profile
    except ImportError:
        normalize_buyer_profile = None

# Google Sheets API scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


class GoogleSheetsDB:
    """
    A database-like interface for Google Sheets.
    Handles buyer profiles, transport modes, and invoice records.
    """
    
    def __init__(self, spreadsheet_id: str = None, credentials_path: str = None):
        """
        Initialize the Google Sheets connection.
        
        Args:
            spreadsheet_id: The ID of the Google Spreadsheet (from URL)
            credentials_path: Path to service account JSON file
        """
        self.spreadsheet_id = spreadsheet_id or os.environ.get('SPREADSHEET_ID')
        self.credentials_path = credentials_path or os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        
        self.client = None
        self.spreadsheet = None
        self._worksheets = {}
        self._cache = {}
        self._cache_time = {}
        self._connect()
    
    def _connect(self):
        """Establish connection to Google Sheets."""
        try:
            if self.credentials_path and os.path.exists(self.credentials_path):
                # Use service account file
                creds = Credentials.from_service_account_file(
                    self.credentials_path, scopes=SCOPES
                )
            else:
                # Use default credentials (for Cloud Run/App Engine)
                from google.auth import default
                creds, _ = default(scopes=SCOPES)
            
            self.client = gspread.authorize(creds)
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            print(f"[OK] Connected to Google Sheets: {self.spreadsheet.title}")
        except Exception as e:
            print(f"[ERROR] Failed to connect to Google Sheets: {e}")
            raise
    
    def _get_or_create_sheet(self, sheet_name: str, headers: List[str]) -> gspread.Worksheet:
        """Get a worksheet or create it if it doesn't exist, caching worksheet reference."""
        if sheet_name in self._worksheets:
            return self._worksheets[sheet_name]
        try:
            worksheet = self.spreadsheet.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            worksheet = self.spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
            worksheet.append_row(headers)
            print(f"✓ Created new sheet: {sheet_name}")
        self._worksheets[sheet_name] = worksheet
        return worksheet
    
    # ===================== BUYER PROFILES =====================
    
    BUYER_HEADERS = ['profile_id', 'buyer_name', 'buyer_details', 'gstin', 'default_tax_type', 'created_at', 'updated_at']
    
    def get_all_buyers(self) -> List[Dict]:
        """Get all buyer profiles with caching and fallback."""
        now = time.time()
        if 'buyers' in self._cache and (now - self._cache_time.get('buyers', 0)) < 300:
            return list(self._cache['buyers'])

        try:
            sheet = self._get_or_create_sheet('Buyers', self.BUYER_HEADERS)
            records = sheet.get_all_records()
            normalized_records = []
            for record in records:
                if record.get('buyer_details'):
                    try:
                        record['buyer_details'] = json.loads(record['buyer_details'])
                    except json.JSONDecodeError:
                        record['buyer_details'] = record['buyer_details'].split('\n')
                else:
                    record['buyer_details'] = []

                if normalize_buyer_profile and record.get('buyer_name'):
                    try:
                        norm = normalize_buyer_profile(record)
                        record.update(norm)
                    except Exception:
                        pass
                normalized_records.append(record)
            records = normalized_records

            # Check if any default buyer from buyer_profiles.json is missing in sheet
            fb_path = os.path.join(os.path.dirname(__file__), 'buyer_profiles.json')
            if os.path.exists(fb_path):
                try:
                    with open(fb_path, 'r', encoding='utf-8') as f:
                        fb_buyers = json.load(f)
                    known_ids = {r.get('profile_id') for r in records}
                    known_names = {r.get('buyer_name', '').strip().lower() for r in records}
                    for fb_b in fb_buyers:
                        if fb_b.get('profile_id') not in known_ids and fb_b.get('buyer_name', '').strip().lower() not in known_names:
                            records.append(fb_b)
                            try:
                                self.save_buyer(fb_b)
                            except Exception:
                                pass
                except Exception:
                    pass

            self._cache['buyers'] = records
            self._cache_time['buyers'] = now
            return records
        except Exception as e:
            print(f"Warning: Failed to fetch buyers from Sheets: {e}")
            if 'buyers' in self._cache:
                return list(self._cache['buyers'])
            fb_path = os.path.join(os.path.dirname(__file__), 'buyer_profiles.json')
            if os.path.exists(fb_path):
                try:
                    with open(fb_path, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except Exception:
                    pass
            return []
    
    def get_buyer(self, profile_id: str) -> Optional[Dict]:
        """Get a specific buyer profile."""
        buyers = self.get_all_buyers()
        return next((b for b in buyers if b.get('profile_id') == profile_id), None)
    
    def save_buyer(self, buyer: Dict) -> bool:
        """Save or update a buyer profile."""
        if normalize_buyer_profile and buyer.get('buyer_name'):
            try:
                buyer = normalize_buyer_profile(buyer)
            except Exception:
                pass

        self._cache.pop('buyers', None)
        sheet = self._get_or_create_sheet('Buyers', self.BUYER_HEADERS)
        
        # Prepare data
        buyer_details_str = json.dumps(buyer.get('buyer_details', []))
        now = datetime.now().isoformat()
        
        # Check if exists
        try:
            cell = sheet.find(buyer['profile_id'], in_column=1)
            # Update existing
            row_num = cell.row
            sheet.update(f'A{row_num}:G{row_num}', [[
                buyer['profile_id'],
                buyer['buyer_name'],
                buyer_details_str,
                buyer.get('gstin', ''),
                buyer.get('default_tax_type', 'IGST'),
                sheet.cell(row_num, 6).value,  # Keep original created_at
                now
            ]])
        except Exception:
            # Insert new
            sheet.append_row([
                buyer['profile_id'],
                buyer['buyer_name'],
                buyer_details_str,
                buyer.get('gstin', ''),
                buyer.get('default_tax_type', 'IGST'),
                now,
                now
            ])
        
        return True
    
    def delete_buyer(self, profile_id: str) -> bool:
        """Delete a buyer profile."""
        sheet = self._get_or_create_sheet('Buyers', self.BUYER_HEADERS)
        try:
            cell = sheet.find(profile_id, in_column=1)
            sheet.delete_rows(cell.row)
            return True
        except Exception:
            return False
    
    # ===================== TRANSPORT MODES =====================
    
    TRANSPORT_HEADERS = ['mode', 'created_at']
    DEFAULT_TRANSPORT_MODES = [
        "Road",
        "By Kolkata Assam Transways",
        "By Gaya Aurangabad Transport",
        "By Bhagirathi Carrying Corporation",
        "By Shree Balaji Roadways",
        "Kolkata Assam Transport LLP",
        "SPS PARRCELL PRIVATE LIMITED",
        "By Vehicle No. WB 11D 4838",
        "By Vehicle No. WB 11E 9369",
    ]
    
    def get_all_transport_modes(self) -> List[str]:
        """Get all transport modes with caching and fallback."""
        now = time.time()
        if 'transports' in self._cache and (now - self._cache_time.get('transports', 0)) < 60:
            return list(self._cache['transports'])
        try:
            sheet = self._get_or_create_sheet('Transport', self.TRANSPORT_HEADERS)
            records = sheet.get_all_records()
            modes = [r['mode'] for r in records if r.get('mode')]
            if not modes:
                modes = list(self.DEFAULT_TRANSPORT_MODES)
                try:
                    for mode in modes:
                        sheet.append_row([mode, datetime.now().isoformat()])
                except Exception:
                    pass
            self._cache['transports'] = modes
            self._cache_time['transports'] = now
            return modes
        except Exception as e:
            print(f"Warning: Failed to fetch transport modes from Sheets: {e}")
            if 'transports' in self._cache:
                return list(self._cache['transports'])
            return list(self.DEFAULT_TRANSPORT_MODES)
    
    def add_transport_mode(self, mode: str) -> bool:
        """Add a new transport mode if it doesn't exist."""
        try:
            sheet = self._get_or_create_sheet('Transport', self.TRANSPORT_HEADERS)
            modes = self.get_all_transport_modes()
            
            # Check if exists (case-insensitive)
            if mode.lower() not in [m.lower() for m in modes]:
                sheet.append_row([mode, datetime.now().isoformat()])
                self._cache.pop('transports', None)
                return True
        except Exception as e:
            print(f"Warning: Failed to add transport mode: {e}")
        return False
    
    # ===================== DISPATCH ADDRESSES =====================
    
    DISPATCH_HEADERS = ['address', 'created_at']
    DEFAULT_DISPATCH_ADDRESSES = [
        "Warehouse Unit 1\nPlot 12, Industrial Estate\nHowrah - 711101, WB (19)",
        "Factory Godown\n45, Phase II Logistics Park\nKolkata - 700001, WB (19)"
    ]

    def get_all_dispatch_addresses(self) -> List[str]:
        """Get all past dispatch addresses with caching and fallback."""
        now = time.time()
        if 'dispatch_addresses' in self._cache and (now - self._cache_time.get('dispatch_addresses', 0)) < 60:
            return list(self._cache['dispatch_addresses'])
        try:
            sheet = self._get_or_create_sheet('Dispatch', self.DISPATCH_HEADERS)
            records = sheet.get_all_records()
            addresses = [r['address'].strip() for r in records if r.get('address') and r['address'].strip()]
            # Ensure defaults are included
            for d in self.DEFAULT_DISPATCH_ADDRESSES:
                if d not in addresses:
                    addresses.append(d)
            self._cache['dispatch_addresses'] = addresses
            self._cache_time['dispatch_addresses'] = now
            return addresses
        except Exception as e:
            print(f"Warning: Failed to fetch dispatch addresses from Sheets: {e}")
            if 'dispatch_addresses' in self._cache:
                return list(self._cache['dispatch_addresses'])
            return list(self.DEFAULT_DISPATCH_ADDRESSES)

    def add_dispatch_address(self, address: str) -> bool:
        """Add a new dispatch address if it doesn't exist."""
        clean_addr = (address or '').strip()
        if not clean_addr:
            return False
        try:
            sheet = self._get_or_create_sheet('Dispatch', self.DISPATCH_HEADERS)
            existing = self.get_all_dispatch_addresses()
            if clean_addr.lower() not in [a.lower() for a in existing]:
                sheet.append_row([clean_addr, datetime.now().isoformat()])
                self._cache.pop('dispatch_addresses', None)
                return True
        except Exception as e:
            print(f"Warning: Failed to add dispatch address: {e}")
        return False

    # ===================== INVOICE RECORDS =====================
    
    LEGACY_INVOICE_HEADERS = [
        'invoice_number', 'invoice_date', 'buyer_name', 'buyer_gstin', 
        'items_json', 'subtotal', 'tax_type', 'tax_amount', 'total_amount',
        'transport_mode', 'file_url', 'pdf_url', 'created_at'
    ]

    INVOICE_HEADERS = [
        'invoice_number', 'invoice_date', 'buyer_name', 'buyer_gstin', 
        'items_json', 'subtotal', 'tax_type', 'display_tax_type', 'tax_rate_igst', 'tax_rate_cgst', 'tax_rate_sgst', 'tax_amount', 'total_amount',
        'transport_mode', 'file_url', 'pdf_url', 'ewaybill_number', 'ewaybill_date', 'ship_from_details', 'ship_to_enabled', 'ship_to_details', 'digitally_signed', 'created_at', 'delivery_charge'
    ]

    def _invoice_row_to_record(self, row: List[str]) -> Dict:
        """Map a raw worksheet row onto the current invoice schema."""
        raw_values = [str(x or '').strip() for x in row]
        # Detect legacy rows: legacy rows had 13 columns (or row length <= 14 where column 7 is numeric tax_amount)
        is_legacy = len(raw_values) <= 14 and (
            len(raw_values) < 8 or raw_values[7].replace('.', '', 1).isdigit() or (len(raw_values) > 6 and raw_values[6] in ('IGST', 'CGST_SGST') and len(raw_values) <= 13)
        )

        if is_legacy:
            values = raw_values + [''] * max(0, len(self.LEGACY_INVOICE_HEADERS) - len(raw_values))
            record = dict(zip(self.LEGACY_INVOICE_HEADERS, values[:len(self.LEGACY_INVOICE_HEADERS)]))
            record['display_tax_type'] = record.get('tax_type', 'IGST')
            record['tax_rate_igst'] = 5.0 if record.get('tax_type') == 'IGST' else 0.0
            record['tax_rate_cgst'] = 2.5 if record.get('tax_type') != 'IGST' else 0.0
            record['tax_rate_sgst'] = 2.5 if record.get('tax_type') != 'IGST' else 0.0
            record['delivery_charge'] = 0.0
            record['ewaybill_number'] = ''
            record['ewaybill_date'] = ''
            record['ship_from_details'] = []
            record['ship_to_enabled'] = False
            record['ship_to_details'] = []
            record['digitally_signed'] = True
        else:
            values = raw_values + [''] * max(0, len(self.INVOICE_HEADERS) - len(raw_values))
            record = dict(zip(self.INVOICE_HEADERS, values[:len(self.INVOICE_HEADERS)]))

        for key in ('subtotal', 'tax_rate_igst', 'tax_rate_cgst', 'tax_rate_sgst', 'tax_amount', 'total_amount', 'delivery_charge'):
            try:
                record[key] = float(record.get(key, 0) or 0)
            except (TypeError, ValueError):
                record[key] = 0.0

        # Ensure ship_from_details and ship_to_details are lists. Accept both
        # pre-parsed lists and JSON-encoded strings stored in Sheets.
        for key in ('ship_from_details', 'ship_to_details'):
            val = record.get(key)
            if isinstance(val, list):
                record[key] = val
                continue
            if isinstance(val, str):
                try:
                    record[key] = json.loads(val) if val.strip() else []
                except json.JSONDecodeError:
                    record[key] = [line.strip() for line in val.splitlines() if line.strip()]
            else:
                record[key] = []

        record['ship_to_enabled'] = str(record.get('ship_to_enabled', '')).strip().lower() in {'1', 'true', 'yes'}
        record['digitally_signed'] = str(record.get('digitally_signed', '')).strip().lower() in {'1', 'true', 'yes'}

        return record
    
    def get_all_invoices(self, limit: Optional[int] = 100) -> List[Dict]:
        """Get recent invoices with caching and fallback."""
        now = time.time()
        if 'invoices' in self._cache and (now - self._cache_time.get('invoices', 0)) < 120:
            records = list(self._cache['invoices'])
            if limit is None or limit <= 0:
                return records
            return records[:limit]

        try:
            sheet = self._get_or_create_sheet('Invoices', self.INVOICE_HEADERS)
            rows = sheet.get_all_values()
            if not rows:
                return []

            records = [self._invoice_row_to_record(row) for row in rows[1:]]

            for record in records:
                if record.get('items_json'):
                    try:
                        record['items'] = json.loads(record['items_json'])
                    except json.JSONDecodeError:
                        record['items'] = []
                else:
                    record['items'] = []

                for key in ('ship_from_details', 'ship_to_details'):
                    val = record.get(key)
                    if isinstance(val, list):
                        continue
                    if isinstance(val, str):
                        try:
                            record[key] = json.loads(val) if val.strip() else []
                        except json.JSONDecodeError:
                            record[key] = [line.strip() for line in val.splitlines() if line.strip()]
                    else:
                        record[key] = []

                record['ship_to_enabled'] = str(record.get('ship_to_enabled', '')).strip().lower() in {'1', 'true', 'yes'}
                record['digitally_signed'] = str(record.get('digitally_signed', '')).strip().lower() in {'1', 'true', 'yes'}
                record['display_tax_type'] = record.get('display_tax_type') or (
                    'CGST_SGST' if str(record.get('tax_type', '')).upper() == 'CGST_SGST' else 'IGST'
                )

            def _sort_inv(x):
                inv_num = str(x.get('invoice_number', '')).strip()
                m = re.search(r'(\d+)', inv_num)
                num = int(m.group(1)) if m else 0
                dt = str(x.get('invoice_date', '') or x.get('created_at', ''))
                return (dt, num)

            records.sort(key=_sort_inv, reverse=True)
            self._cache['invoices'] = records
            self._cache_time['invoices'] = now
            if limit is None or limit <= 0:
                return records
            return records[:limit]
        except Exception as e:
            print(f"Warning: Failed to fetch invoices from Sheets: {e}")
            if 'invoices' in self._cache:
                records = list(self._cache['invoices'])
                if limit is None or limit <= 0:
                    return records
                return records[:limit]
            return []
    
    def save_invoice(self, invoice: Dict) -> bool:
        """Save an invoice record."""
        sheet = self._get_or_create_sheet('Invoices', self.INVOICE_HEADERS)
        
        items_json = json.dumps(invoice.get('items', []))
        now = datetime.now().isoformat()
        
        sheet.append_row([
            invoice.get('invoice_number', ''),
            invoice.get('invoice_date', ''),
            invoice.get('buyer_name', ''),
            invoice.get('buyer_gstin', ''),
            items_json,
            invoice.get('subtotal', 0),
            invoice.get('tax_type', 'IGST'),
            invoice.get('display_tax_type', 'IGST'),
            invoice.get('tax_rate_igst', 5.0),
            invoice.get('tax_rate_cgst', 2.5),
            invoice.get('tax_rate_sgst', 2.5),
            invoice.get('tax_amount', 0),
            invoice.get('total_amount', 0),
            invoice.get('transport_mode', ''),
            invoice.get('file_url', ''),
            invoice.get('pdf_url', ''),
            invoice.get('ewaybill_number', ''),
            invoice.get('ewaybill_date', ''),
            json.dumps(invoice.get('ship_from_details', [])),
            str(bool(invoice.get('ship_to_enabled', False))).lower(),
            json.dumps(invoice.get('ship_to_details', [])),
            str(bool(invoice.get('digitally_signed', False))).lower(),
            now,
            invoice.get('delivery_charge', 0.0)
        ])
        self._cache.pop('invoices', None)
        return True
    
    def get_invoice(self, invoice_number: str) -> Optional[Dict]:
        """Get a specific invoice by number with tolerant matching."""
        target = (invoice_number or '').strip()
        if not target:
            return None

        # Search across full history
        invoices = self.get_all_invoices(limit=None)
        # 1. Exact match
        for i in invoices:
            if (i.get('invoice_number') or '').strip() == target:
                return i

        # 2. Normalized prefix match (e.g. 61 vs 061, with or without FY)
        target_norm = target.replace('/', '-').replace(' ', '').lower()
        target_m = re.search(r'(\d+)', target)
        target_num = int(target_m.group(1)) if target_m else None

        for i in invoices:
            inv_str = str(i.get('invoice_number') or '').strip()
            inv_norm = inv_str.replace('/', '-').replace(' ', '').lower()
            if inv_norm == target_norm:
                return i
            if target_num is not None:
                inv_m = re.search(r'(\d+)', inv_str)
                if inv_m and int(inv_m.group(1)) == target_num:
                    return i
        return None
    
    def get_last_invoice_number(self) -> Optional[str]:
        """Get the highest sequential invoice number for current FY."""
        invoices = self.get_all_invoices(limit=2000)
        max_num = 0
        best_inv = None
        for inv in invoices:
            inv_str = str(inv.get('invoice_number', '')).strip()
            if '2026-27' in inv_str or '26-27' in inv_str:
                m = re.search(r'(\d+)', inv_str)
                if m:
                    val = int(m.group(1))
                    # Skip outlier test bill 99
                    if val != 99 and val > max_num:
                        max_num = val
                        best_inv = inv_str
        return best_inv if best_inv else (invoices[0].get('invoice_number') if invoices else None)


# ===================== HELPER FUNCTIONS =====================

def init_sheets_db() -> GoogleSheetsDB:
    """Initialize and return the Google Sheets database connection."""
    spreadsheet_id = os.environ.get('SPREADSHEET_ID')
    if not spreadsheet_id:
        raise ValueError("SPREADSHEET_ID environment variable is not set")
    
    return GoogleSheetsDB(spreadsheet_id=spreadsheet_id)


# For testing locally
if __name__ == '__main__':
    import os
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'service-account.json'
    os.environ['SPREADSHEET_ID'] = 'YOUR_SPREADSHEET_ID'
    
    db = init_sheets_db()
    print("Buyers:", db.get_all_buyers())
    print("Transport:", db.get_all_transport_modes())
    print("Invoices:", db.get_all_invoices(limit=5))
