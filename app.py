"""
Shakambhari Enterprises Invoice Generator - Main Application

A Flask-based invoice generation system for GST billing with features:
- Buyer profile management with search/filter
- Multiple line items support
- Transport mode management (type or select)
- Load and edit existing invoices
- Excel generation with formulas
- PDF conversion (Windows only)
"""

import os
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify, make_response
from datetime import datetime
import uuid
from num2words import num2words
from typing import Any, List, Dict, Optional
# The old helper `copy1.copy_excel_with_formatting` was removed; use the
# cloud.app_cloud.generate_invoice_excel exporter when available.
try:
    from cloud.app_cloud import generate_invoice_excel, _extract_invoice_data_from_xlsx_bytes
except Exception:
    generate_invoice_excel = None


def copy_excel_with_formatting(template_path: str, dest_path: str, config_data: dict):
    """Compatibility wrapper used by existing code paths.
    It builds a minimal invoice payload and delegates to
    `cloud.app_cloud.generate_invoice_excel` when present, writing bytes
    to `dest_path`.
    """
    if not generate_invoice_excel:
        raise RuntimeError('Excel generation helper not available (generate_invoice_excel missing)')

    invoice_payload = {
        'invoice_number': config_data.get('invoice_number', ''),
        'invoice_date_display': config_data.get('invoice_date', ''),
        'buyer_details': config_data.get('buyer_details', []),
        'transport_mode': config_data.get('mode_of_transport', ''),
        'items': config_data.get('items', []),
        'delivery_charge': config_data.get('delivery_charge', 0.0),
        'tax_type': config_data.get('tax_type', 'IGST'),
        # Optional fields that may be present elsewhere in the app
        'ewaybill_number': config_data.get('ewaybill_number', ''),
        'ewaybill_date_display': config_data.get('ewaybill_date_display', ''),
        'ship_from_details': config_data.get('ship_from_details', []),
        'ship_to_enabled': config_data.get('ship_to_enabled', False),
        'ship_to_details': config_data.get('ship_to_details', []),
        'digitally_signed': config_data.get('digitally_signed', False),
        'display_tax_type': config_data.get('tax_type', 'IGST')
    }

    excel_bytes = generate_invoice_excel(invoice_payload)
    with open(dest_path, 'wb') as f:
        f.write(excel_bytes)
from config import (
    BUYER_PROFILES_JSON, TRANSPORT_MODES_JSON, OUTPUT_DIR, 
    PDF_OUTPUT_DIR, TEMPLATE_EXCEL_FILE, ensure_dirs, BASE_DIR
)
from settings_manager import get_settings, update_settings, DEFAULT_SETTINGS

# Attempt to import win32com.client for PDF conversion
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("WARNING: pywin32 library not found. PDF conversion will be skipped.")

# Optional openpyxl for reading invoices
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("WARNING: openpyxl not found. Loading old invoices will be limited.")

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or os.urandom(24).hex()
app.config['SESSION_COOKIE_NAME'] = '__session'
ensure_dirs()

BACKUP_DIR = os.path.join(BASE_DIR, "_backups")
os.makedirs(BACKUP_DIR, exist_ok=True)


# ===================== UTILITY FUNCTIONS =====================

def backup_json(path: str) -> None:
    """Create a timestamped backup of a JSON file."""
    if not os.path.isfile(path):
        return
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(BACKUP_DIR, f"{os.path.basename(path)}.{stamp}.bak")
        with open(path, 'rb') as src, open(backup_file, 'wb') as dst:
            dst.write(src.read())
    except Exception as e:
        print(f"WARNING: Backup failed for {path}: {e}")


def load_data(json_path: str) -> List:
    """Load JSON data from a file."""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON from {json_path}: {e}")
        return []


def save_data(json_path: str, data: Any) -> bool:
    """Save data to a JSON file with backup."""
    try:
        backup_json(json_path)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except IOError as e:
        print(f"Error saving data to {json_path}: {e}")
        return False


def _financial_year_suffix(today: datetime = None) -> str:
    """Get the financial year suffix like /2025-26."""
    today = today or datetime.now()
    year = today.year
    if today.month >= 4:
        start = year
        end = year + 1
    else:
        start = year - 1
        end = year
    return f"/{start}-{str(end)[-2:]}"


def safe_float(value: Any, default: float = 0.0) -> float:
    """Safely convert a value to float with a fallback."""
    try:
        if value is None or value == '':
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def round_half_up(value: float) -> int:
    """Round to nearest integer with .5 always rounding up."""
    return int(Decimal(str(value)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def unique_filename(directory: str, base_name: str, extension: str) -> str:
    """Generate a unique filename in a directory to avoid accidental overwrite."""
    candidate = f"{base_name}{extension}"
    counter = 2
    while os.path.exists(os.path.join(directory, candidate)):
        candidate = f"{base_name}__{counter}{extension}"
        counter += 1
    return candidate


def is_new_invoice_layout(sheet) -> bool:
    """Detect if the sheet uses the new layout that includes delivery charge row."""
    row_label = str(sheet['C30'].value or '').strip().lower()
    return 'delivery' in row_label


def next_invoice_number(existing_files: List[str]) -> str:
    """Calculate the next invoice number based on existing files of current FY only."""
    max_num = 0
    fy_token = _financial_year_suffix().lstrip('/').replace('-', '_')
    pattern = re.compile(rf"^Invoice_(\d+)_{re.escape(fy_token)}(?:_|$)", re.IGNORECASE)
    for fname in existing_files:
        m = pattern.search(fname)
        if m:
            try:
                num = int(m.group(1))
                if num > max_num:
                    max_num = num
            except ValueError:
                continue
    return f"{max_num + 1:03d}{_financial_year_suffix()}"


def suggest_next_invoice_number() -> str:
    """Get the suggested next invoice number."""
    try:
        files = os.listdir(OUTPUT_DIR)
    except FileNotFoundError:
        files = []
    return next_invoice_number(files)


TRANSPORT_PREFIX_VARIANTS = [
    'mode of transport:', 'mode of transports:', 
    'mode of transport', 'mode of transports'
]


def extract_transport_core(raw: str) -> str:
    """Extract the core transport value without prefix."""
    if not raw:
        return ''
    val = raw.strip()
    low = val.lower()
    for prefix in TRANSPORT_PREFIX_VARIANTS:
        if low.startswith(prefix):
            val = val[len(prefix):].strip(' -:')
            break
    return val.strip()


def normalize_transport_mode(raw: str) -> str:
    """Normalize transport mode to a canonical format."""
    core = extract_transport_core(raw)
    if not core:
        return ''
    return f"Mode of Transport: {core}"


def save_new_transport_mode(transport_value: str) -> bool:
    """Save a new transport mode to the JSON file if it doesn't exist."""
    if not transport_value:
        return False
    
    core_value = extract_transport_core(transport_value)
    if not core_value:
        return False
    
    transport_modes = load_data(TRANSPORT_MODES_JSON)
    
    # Check if this transport mode already exists (case-insensitive)
    existing_cores = set()
    for mode in transport_modes:
        existing_cores.add(extract_transport_core(mode).lower())
    
    if core_value.lower() not in existing_cores:
        transport_modes.append(f"Mode of Transport: {core_value}")
        save_data(TRANSPORT_MODES_JSON, transport_modes)
        print(f"Saved new transport mode: {core_value}")
        return True
    return False


def convert_excel_to_pdf(excel_filepath: str, pdf_filepath: str) -> bool:
    """Convert an Excel file to PDF using Excel COM automation."""
    if not WIN32COM_AVAILABLE:
        print(f"Skipping PDF conversion - pywin32 not available")
        return False
    
    try:
        import pythoncom
    except Exception:
        print("Skipping PDF conversion - pythoncom not available")
        return False

    excel = None
    workbook = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        if not os.path.exists(excel_filepath):
            print(f"Error: Excel file not found at {excel_filepath}")
            return False
        
        workbook = excel.Workbooks.Open(excel_filepath)
        os.makedirs(os.path.dirname(pdf_filepath), exist_ok=True)
        workbook.ExportAsFixedFormat(
            0, pdf_filepath, 
            Quality=0, 
            IncludeDocProperties=True, 
            IgnorePrintAreas=False, 
            OpenAfterPublish=False
        )
        print(f"PDF created: {pdf_filepath}")
        return True
    except Exception as e:
        print(f"Error converting to PDF: {e}")
        return False
    finally:
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()


def get_generated_invoices() -> List[Dict]:
    """Get list of all generated invoices with metadata."""
    invoices = []
    try:
        files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith('.xlsx') and f.startswith('Invoice_')]
        files.sort(key=lambda f: os.path.getmtime(os.path.join(OUTPUT_DIR, f)), reverse=True)
        for fname in files:
            if fname.endswith('.xlsx') and fname.startswith('Invoice_'):
                filepath = os.path.join(OUTPUT_DIR, fname)
                try:
                    mtime = os.path.getmtime(filepath)
                    date_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
                    
                    # Parse invoice number and buyer from filename
                    invoice_num = ''
                    buyer_name = ''
                    stem = fname.replace('.xlsx', '')
                    m = re.match(r'^Invoice_(\d+)_((\d{4})_(\d{2}))_(.+?)(?:__\d+)?$', stem, re.IGNORECASE)
                    if m:
                        raw_num = m.group(1)
                        year_start = m.group(3)
                        year_end = m.group(4)
                        buyer_name = m.group(5).replace('_', ' ')
                        invoice_num = f"{int(raw_num):03d}/{year_start}-{year_end}"
                    else:
                        parts = stem.split('_')
                        invoice_num = parts[1] if len(parts) > 1 else ''
                        buyer_name = ' '.join(parts[2:]) if len(parts) > 2 else ''
                        buyer_name = buyer_name.replace('_', ' ')
                    
                    invoice_info = {
                        'filename': fname,
                        'filepath': filepath,
                        'invoice_number': invoice_num,
                        'buyer_name': buyer_name,
                        'modified_date': date_str,
                        'total_amount': '',
                        'items_count': 0,
                        'tax_type': '',
                        'transport_mode': ''
                    }
                    
                    # Try to extract additional details from the Excel file
                    if OPENPYXL_AVAILABLE:
                        try:
                            wb = openpyxl.load_workbook(filepath, data_only=True)
                            sheet = wb.active
                            new_layout = is_new_invoice_layout(sheet)
                            
                            # Read final total from the active layout.
                            total_candidates = ['I36', 'I35', 'I34'] if new_layout else ['I35', 'I34', 'I33']
                            total = None
                            for total_cell in total_candidates:
                                cell_val = sheet[total_cell].value
                                if isinstance(cell_val, (int, float)):
                                    total = cell_val
                                    break
                            if isinstance(total, (int, float)):
                                invoice_info['total_amount'] = f"{total:,.2f}"
                            
                            # Count items (rows 18-27)
                            items_count = 0
                            for row in range(18, 28):
                                desc = sheet[f'A{row}'].value
                                qty = sheet[f'F{row}'].value
                                if desc or qty:
                                    items_count += 1
                            invoice_info['items_count'] = items_count
                            
                            # Get tax type
                            if new_layout:
                                rate31 = str(sheet['E31'].value or '').strip()
                                rate32 = str(sheet['E32'].value or '').strip()
                                val31 = safe_float(sheet['I31'].value, 0.0)
                                val32 = safe_float(sheet['I32'].value, 0.0)
                                if rate31.startswith('2.50') or rate32.startswith('2.50') or (val31 > 0 and not rate31.startswith('5.00')):
                                    invoice_info['tax_type'] = 'CGST+SGST'
                                else:
                                    invoice_info['tax_type'] = 'IGST'
                            else:
                                cgst_val = safe_float(sheet['I31'].value, 0.0)
                                cgst_rate = str(sheet['E31'].value or '').strip()
                                if cgst_val > 0 or cgst_rate.startswith('2.50'):
                                    invoice_info['tax_type'] = 'CGST+SGST'
                                else:
                                    invoice_info['tax_type'] = 'IGST'
                            
                            # Get transport mode
                            transport = sheet['E10'].value or ''
                            transport_core = extract_transport_core(str(transport))
                            invoice_info['transport_mode'] = transport_core
                            
                            wb.close()
                        except Exception as e:
                            # Silently fail - we still have basic info from filename
                            pass
                    
                    invoices.append(invoice_info)
                except Exception:
                    continue
    except FileNotFoundError:
        pass
    return invoices


def extract_invoice_data(filepath: str) -> Optional[Dict]:
    """Extract data from an existing invoice Excel file."""
    if not os.path.exists(filepath):
        return None
    try:
        with open(filepath, 'rb') as f:
            file_bytes = f.read()
        return _extract_invoice_data_from_xlsx_bytes(file_bytes, os.path.basename(filepath))
    except Exception as e:
        print(f"Error extracting invoice data: {e}")
        import traceback
        traceback.print_exc()
        return None


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
    buyer_profiles = load_data(BUYER_PROFILES_JSON)
    transport_modes = load_data(TRANSPORT_MODES_JSON)
    today_date = datetime.now().strftime('%Y-%m-%d')
    
    # Filter valid profiles and sort by name
    valid_buyer_profiles = [p for p in buyer_profiles if p.get('profile_id') and p.get('buyer_name')]
    valid_buyer_profiles.sort(key=lambda p: p.get('buyer_name', '').lower())
    
    # Normalize and deduplicate transport modes
    transport_cores = []
    seen = set()
    for mode in transport_modes:
        core = extract_transport_core(mode)
        if core and core.lower() not in seen:
            seen.add(core.lower())
            transport_cores.append(core)
    transport_cores.sort()
    
    suggestion = suggest_next_invoice_number()
    recent_invoices = get_generated_invoices()[:50]  # Increased for modal
    
    # Check if loading a specific invoice
    load_filename = request.args.get('load', '')
    preload_invoice = None
    if load_filename:
        filepath = os.path.join(OUTPUT_DIR, load_filename)
        if os.path.exists(filepath):
            preload_invoice = extract_invoice_data(filepath)
            if preload_invoice:
                preload_invoice['filename'] = load_filename
    
    return render_template('index.html', 
                          buyer_profiles=valid_buyer_profiles, 
                          transport_modes=transport_cores, 
                          today_date=today_date, 
                          suggested_invoice_number=suggestion,
                          recent_invoices=recent_invoices,
                          preload_invoice=preload_invoice,
                          app_settings=get_settings())


@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    """Generate an invoice from form data."""
    try:
        # Core form fields
        buyer_profile_id = request.form.get('buyer_profile_id')
        if not buyer_profile_id:
            flash("Please select a buyer profile.", "error")
            return redirect(url_for('index'))
        
        raw_invoice_number = request.form.get('invoice_number', '').strip()
        invoice_date_str = request.form.get('invoice_date', '')
        
        # Parse date
        try:
            dt_object = datetime.strptime(invoice_date_str, '%Y-%m-%d')
        except ValueError:
            flash("Invalid date format. Please use YYYY-MM-DD.", "error")
            return redirect(url_for('index'))
        
        formatted_invoice_date = dt_object.strftime('%d/%m/%Y')
        
        # Invoice display strings
        invoice_number_for_filename = raw_invoice_number
        excel_invoice_number_display = f"INVOICE No. {invoice_number_for_filename}" if invoice_number_for_filename else ""
        excel_invoice_date_display = f"Date : {formatted_invoice_date}" if formatted_invoice_date else ""
        
        # Transport mode - accept typed input and save if new
        transport_mode_input = request.form.get('transport_mode', '').strip()
        transport_mode = normalize_transport_mode(transport_mode_input)

        # Delivery charge is invoice-level and taxable.
        delivery_charge_raw = request.form.get('delivery_charge', '0').strip()
        delivery_charge = safe_float(delivery_charge_raw, 0.0)
        if delivery_charge < 0:
            flash("Delivery charge cannot be negative.", "error")
            return redirect(url_for('index'))
        
        if transport_mode_input:
            save_new_transport_mode(transport_mode_input)
        
        # Process multiple items
        items = []
        item_descriptions = request.form.getlist('item_description[]')
        item_bags = request.form.getlist('item_bags[]')
        item_units = request.form.getlist('item_unit[]')
        item_hsns = request.form.getlist('item_hsn[]')
        item_quantities = request.form.getlist('item_quantity[]')
        item_rates = request.form.getlist('item_rate[]')
        
        if item_descriptions:
            for i in range(len(item_descriptions)):
                desc = item_descriptions[i].strip() if i < len(item_descriptions) else ''
                bags = item_bags[i].strip() if i < len(item_bags) else ''
                unit_val = item_units[i].strip() if i < len(item_units) and item_units[i].strip() else 'Bags'
                hsn = item_hsns[i].strip() if i < len(item_hsns) else ''
                qty = float(item_quantities[i]) if i < len(item_quantities) and item_quantities[i] else 0
                rt = float(item_rates[i]) if i < len(item_rates) and item_rates[i] else 0
                
                if desc or qty or rt:
                    clean_bag = re.sub(r'(?i)\s*(?:bags?|bundles?|boxes?|pkts?|cartons?)', '', bags).strip()
                    desc = re.sub(r'\s*\(\s*[\d\.]+\s*(?:[A-Za-z]+|\s*)+\)', '', desc, flags=re.IGNORECASE).strip()
                    full_desc = desc
                    if clean_bag:
                        try:
                            bags_num = float(clean_bag)
                            bags_str = f"{int(bags_num)}" if bags_num.is_integer() else f"{bags_num}"
                            full_desc += f" ({bags_str} {unit_val})"
                        except ValueError:
                            full_desc += f" ({clean_bag} {unit_val})"
                    
                    items.append({
                        'description': full_desc,
                        'bags': clean_bag,
                        'unit': unit_val,
                        'hsn': hsn,
                        'quantity': qty,
                        'rate': rt
                    })
        else:
            # Backward compatibility - single item
            base_desc = request.form.get('item_base_description', '1. Aluminium Utensils').strip()
            bags = request.form.get('item_description_bags', '').strip()
            hsn = request.form.get('item_hsn', '').strip()
            quantity = float(request.form.get('quantity', 0) or 0)
            rate = float(request.form.get('rate', 0) or 0)
            
            description = base_desc
            if bags:
                try:
                    bags_int = int(float(bags))
                    if bags_int > 0:
                        description += f" ({bags_int} Bags)"
                except ValueError:
                    description += f" ({bags} Bags)"
            
            items.append({
                'description': description,
                'hsn': hsn,
                'quantity': quantity,
                'rate': rate
            })
        
        if not items:
            flash("Please add at least one item.", "error")
            return redirect(url_for('index'))
        
        # Buyer profile lookup
        buyer_profiles = load_data(BUYER_PROFILES_JSON)
        selected_profile = next((p for p in buyer_profiles if p.get('profile_id') == buyer_profile_id), None)
        if not selected_profile:
            flash("Selected buyer profile not found.", "error")
            return redirect(url_for('index'))
        
        # Tax type resolution
        tax_type_override = request.form.get('tax_type_override')
        if tax_type_override and tax_type_override != "PROFILE_DEFAULT":
            final_tax_type = tax_type_override
        else:
            final_tax_type = selected_profile.get('default_tax_type', 'IGST')
        
        # Build config for Excel generation
        config_data = {
            "buyer_details": selected_profile.get('buyer_details', []),
            "mode_of_transport": transport_mode,
            "items": items,
            "item_details": items[0] if items else {"description": "", "quantity": 0, "rate": 0},
            "delivery_charge": delivery_charge,
            "tax_type": final_tax_type,
            "invoice_number": excel_invoice_number_display,
            "invoice_date": excel_invoice_date_display
        }
        
        # Generate filenames
        safe_invoice_number = ''.join(c if c.isalnum() else '_' for c in invoice_number_for_filename)
        safe_buyer_name = ''.join(c if c.isalnum() else '_' for c in selected_profile.get('buyer_name', 'Unknown'))
        
        excel_filename_base = f"Invoice_{safe_invoice_number}_{safe_buyer_name}" if safe_invoice_number else f"Invoice_{safe_buyer_name}"
        excel_output_filename = unique_filename(OUTPUT_DIR, excel_filename_base, '.xlsx')
        excel_destination_filepath = os.path.join(OUTPUT_DIR, excel_output_filename)

        if excel_output_filename != f"{excel_filename_base}.xlsx":
            flash(f"Same invoice number already existed. Saved as {excel_output_filename}.", "warning")
        
        if not TEMPLATE_EXCEL_FILE:
            flash("No Excel template found. Place a template .xlsx inside the 'GST Invoices' folder.", "error")
            return redirect(url_for('index'))
        
        # Generate Excel
        copy_excel_with_formatting(TEMPLATE_EXCEL_FILE, excel_destination_filepath, config_data)
        
        # PDF conversion
        pdf_base_name = os.path.splitext(excel_output_filename)[0]
        pdf_output_filename = unique_filename(PDF_OUTPUT_DIR, pdf_base_name, '.pdf')
        pdf_destination_filepath = os.path.join(PDF_OUTPUT_DIR, pdf_output_filename)
        
        if WIN32COM_AVAILABLE and convert_excel_to_pdf(excel_destination_filepath, pdf_destination_filepath):
            flash(f"Invoice {excel_output_filename} generated with PDF!", "success")
            return redirect(url_for('success_pdf', filename=pdf_output_filename))
        elif WIN32COM_AVAILABLE:
            flash(f"Invoice {excel_output_filename} generated, but PDF conversion failed.", "warning")
            return redirect(url_for('success', filename=excel_output_filename))
        else:
            flash(f"Invoice {excel_output_filename} generated successfully!", "success")
            return redirect(url_for('success', filename=excel_output_filename))
        
    except Exception as e:
        flash(f"Error generating invoice: {e}", "error")
        print(f"Exception in generate_invoice: {e}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('index'))


@app.route('/calculate_preview', methods=['POST'])
def calculate_preview_route():
    """Calculate invoice totals for live preview."""
    data = request.json
    try:
        items = data.get('items', [])
        if not items:
            quantity = float(data.get('quantity', 0))
            rate = float(data.get('rate', 0))
            items = [{'quantity': quantity, 'rate': rate}]
        
        tax_type = data.get('tax_type', 'IGST')
        delivery_charge = safe_float(data.get('delivery_charge', 0), 0.0)
        if delivery_charge < 0:
            delivery_charge = 0.0
        
        # Calculate totals
        item_amounts = []
        subtotal = 0
        for item in items:
            qty = float(item.get('quantity', 0))
            rt = float(item.get('rate', 0))
            amount = qty * rt
            item_amounts.append(amount)
            subtotal += amount
        
        taxable_amount = subtotal + delivery_charge

        # Tax calculations
        igst_amount = 0
        cgst_amount = 0
        sgst_amount = 0
        
        if tax_type == "IGST":
            igst_amount = taxable_amount * 0.05
        elif tax_type == "CGST_SGST":
            cgst_amount = taxable_amount * 0.025
            sgst_amount = taxable_amount * 0.025
        
        total_before_round = taxable_amount + igst_amount + cgst_amount + sgst_amount
        rounded_total = round_half_up(total_before_round)
        round_off = rounded_total - total_before_round
        
        # Amount in words
        if rounded_total > 0:
            amount_words = num2words(int(rounded_total), lang='en_IN')
            amount_words = amount_words.replace('-', ' ').replace(',', '').title() + " Only"
        else:
            amount_words = "Zero Only"
        
        return jsonify({
            "item_amounts": [f"{a:.2f}" for a in item_amounts],
            "item_amount": f"{item_amounts[0]:.2f}" if item_amounts else "0.00",
            "subtotal": f"{subtotal:.2f}",
            "delivery_charge": f"{delivery_charge:.2f}",
            "taxable_amount": f"{taxable_amount:.2f}",
            "igst_amount": f"{igst_amount:.2f}",
            "cgst_amount": f"{cgst_amount:.2f}",
            "sgst_amount": f"{sgst_amount:.2f}",
            "total_before_round_off": f"{total_before_round:.2f}",
            "round_off_value": f"{round_off:.2f}",
            "rounded_total": f"{rounded_total:.2f}",
            "amount_in_words": amount_words
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route('/api/invoice/<path:invoice_number>')
def api_get_invoice(invoice_number):
    """Compatibility route to fetch an invoice by number or filename."""
    clean_target = (invoice_number or '').replace('/', '-').replace('\\', '-').strip()
    m_num = re.search(r'(\d+)', clean_target)
    target_int = int(m_num.group(1)) if m_num else None

    # Search generated invoices folder
    if os.path.exists(OUTPUT_DIR):
        for fname in os.listdir(OUTPUT_DIR):
            if not fname.endswith('.xlsx'):
                continue
            match = False
            if clean_target and clean_target.lower() in fname.lower():
                match = True
            elif target_int is not None:
                fm = re.search(r'Invoice_0*(\d+)', fname)
                if fm and int(fm.group(1)) == target_int:
                    match = True
            if match:
                data = extract_invoice_data(os.path.join(OUTPUT_DIR, fname))
                if data:
                    return jsonify({'success': True, 'invoice': data, **data})

    return jsonify({"error": "Invoice not found", "success": False}), 404


@app.route('/api/load_invoice/<filename>')
def api_load_invoice(filename):
    """Load invoice data from an existing Excel file."""
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    if not os.path.exists(filepath):
        return jsonify({"error": "Invoice file not found"}), 404
    
    data = extract_invoice_data(filepath)
    if data:
        return jsonify(data)
    else:
        return jsonify({"error": "Failed to extract invoice data"}), 500


@app.route('/api/invoices')
def api_list_invoices():
    """List all generated invoices."""
    return jsonify(get_generated_invoices())


@app.route('/api/next_invoice_number')
def api_next_invoice_number():
    """Get the next suggested invoice number."""
    return jsonify({"next_invoice_number": suggest_next_invoice_number()})


@app.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    """Retrieve or update dynamic application settings."""
    if request.method == 'POST':
        updates = request.get_json(silent=True) or {}
        new_settings = update_settings(updates)
        return jsonify({'success': True, 'settings': new_settings})
    return jsonify({'success': True, 'settings': get_settings()})


@app.route('/api/ai/parse-bill', methods=['POST'])
def api_ai_parse_bill_local():
    """Delegate AI parse bill to cloud module implementation if available."""
    try:
        from cloud.app_cloud import api_ai_parse_bill
        return api_ai_parse_bill()
    except Exception as e:
        return jsonify({'success': False, 'error': f'AI parse error: {e}'}), 500


@app.route('/api/ai/create-buyer', methods=['POST'])
def api_ai_create_buyer_local():
    """Delegate AI buyer creation to cloud module implementation if available."""
    try:
        from cloud.app_cloud import api_ai_create_buyer
        return api_ai_create_buyer()
    except Exception as e:
        return jsonify({'success': False, 'error': f'AI create buyer error: {e}'}), 500


@app.route('/download/xlsx/<path:filename>')
@app.route('/download_xlsx/<path:filename>')
def download_xlsx(filename):
    """Download an Excel invoice."""
    try:
        from cloud.app_cloud import download_xlsx as dl_x
        return dl_x(filename)
    except Exception as e:
        return redirect(url_for('index'))


@app.route('/download/pdf/<path:filename>')
@app.route('/download_pdf/<path:filename>')
def download_pdf(filename):
    """Download a PDF invoice."""
    try:
        from cloud.app_cloud import download_pdf as dl_p
        return dl_p(filename)
    except Exception as e:
        return redirect(url_for('index'))


# ===================== PROFILE MANAGEMENT =====================

@app.route('/list_profiles')
def list_profiles():
    """List all buyer profiles."""
    buyer_profiles = load_data(BUYER_PROFILES_JSON)
    valid_profiles = [p for p in buyer_profiles if p.get('profile_id') and p.get('buyer_name')]
    valid_profiles.sort(key=lambda p: p.get('buyer_name', '').lower())
    return render_template('list_profiles.html', profiles=valid_profiles)


@app.route('/manage_profile', methods=['GET', 'POST'])
@app.route('/manage_profile/<profile_id>', methods=['GET', 'POST'])
def manage_profile(profile_id=None):
    """Create or edit a buyer profile."""
    buyer_profiles = load_data(BUYER_PROFILES_JSON)
    profile_to_edit = None
    is_new_profile = False
    
    if profile_id:
        profile_to_edit = next((p for p in buyer_profiles if p.get('profile_id') == profile_id), None)
        if not profile_to_edit:
            flash(f"Profile not found.", "error")
            return redirect(url_for('list_profiles'))
    else:
        is_new_profile = True
        profile_to_edit = {
            "buyer_name": "",
            "buyer_details": [],
            "gstin": "",
            "default_tax_type": "IGST",
            "profile_id": ""
        }
    
    if request.method == 'POST':
        buyer_name = request.form.get('buyer_name', '').strip()
        buyer_details_str = request.form.get('buyer_details_textarea', '')
        buyer_details = [line.strip() for line in buyer_details_str.split('\n') if line.strip()]
        gstin = request.form.get('gstin', '').strip().upper()
        default_tax_type = request.form.get('default_tax_type', 'IGST')
        
        if not buyer_name:
            flash("Buyer Name is required.", "error")
            profile_data = {
                'buyer_name': buyer_name,
                'buyer_details_textarea': buyer_details_str,
                'gstin': gstin,
                'default_tax_type': default_tax_type,
                'profile_id': profile_id or ''
            }
            return render_template('profile_form.html', profile=profile_data, 
                                 is_new_profile=is_new_profile, profile_id=profile_id)
        
        if is_new_profile:
            # Generate profile ID
            if gstin:
                new_profile_id = gstin
            else:
                safe_name = ''.join(c if c.isalnum() else '_' for c in buyer_name)
                new_profile_id = f"{safe_name}_{uuid.uuid4().hex[:8]}"
            
            # Check for duplicates
            if any(p.get('profile_id') == new_profile_id for p in buyer_profiles):
                flash(f"A profile with this ID already exists.", "error")
                profile_data = {
                    'buyer_name': buyer_name,
                    'buyer_details_textarea': buyer_details_str,
                    'gstin': gstin,
                    'default_tax_type': default_tax_type,
                    'profile_id': ''
                }
                return render_template('profile_form.html', profile=profile_data, 
                                     is_new_profile=True, profile_id=None)
            
            new_profile = {
                "profile_id": new_profile_id,
                "buyer_name": buyer_name,
                "buyer_details": buyer_details,
                "gstin": gstin,
                "default_tax_type": default_tax_type
            }
            buyer_profiles.append(new_profile)
            flash(f"Profile '{buyer_name}' created successfully!", "success")
        else:
            profile_to_update = next((p for p in buyer_profiles if p.get('profile_id') == profile_id), None)
            if profile_to_update:
                profile_to_update['buyer_name'] = buyer_name
                profile_to_update['buyer_details'] = buyer_details
                profile_to_update['gstin'] = gstin
                profile_to_update['default_tax_type'] = default_tax_type
                flash(f"Profile '{buyer_name}' updated successfully!", "success")
            else:
                flash("Error: Profile not found for update.", "error")
                return redirect(url_for('list_profiles'))
        
        if save_data(BUYER_PROFILES_JSON, buyer_profiles):
            return redirect(url_for('list_profiles'))
        else:
            flash("Error saving profile.", "error")
            return redirect(url_for('list_profiles'))
    
    # GET request
    if profile_to_edit and isinstance(profile_to_edit.get('buyer_details'), list):
        profile_to_edit['buyer_details_textarea'] = '\n'.join(profile_to_edit['buyer_details'])
    elif profile_to_edit:
        profile_to_edit['buyer_details_textarea'] = ''
    
    return render_template('profile_form.html', profile=profile_to_edit, 
                          is_new_profile=is_new_profile, profile_id=profile_id)


@app.route('/delete_profile/<profile_id>', methods=['POST'])
def delete_profile(profile_id):
    """Delete a buyer profile."""
    buyer_profiles = load_data(BUYER_PROFILES_JSON)
    
    profile_to_delete = next((p for p in buyer_profiles if p.get('profile_id') == profile_id), None)
    if not profile_to_delete:
        flash("Profile not found.", "error")
        return redirect(url_for('list_profiles'))
    
    buyer_name = profile_to_delete.get('buyer_name', 'Unknown')
    buyer_profiles = [p for p in buyer_profiles if p.get('profile_id') != profile_id]
    
    if save_data(BUYER_PROFILES_JSON, buyer_profiles):
        flash(f"Profile '{buyer_name}' deleted.", "success")
    else:
        flash("Error deleting profile.", "error")
    
    return redirect(url_for('list_profiles'))


@app.route('/cleanup_profiles', methods=['POST'])
def cleanup_profiles():
    """Remove duplicate and invalid profiles."""
    buyer_profiles = load_data(BUYER_PROFILES_JSON)
    original_count = len(buyer_profiles)
    
    # Remove invalid profiles
    valid_profiles = [p for p in buyer_profiles if p.get('profile_id') and p.get('buyer_name')]
    
    # Remove duplicates based on profile_id
    seen_ids = set()
    unique_profiles = []
    for p in valid_profiles:
        pid = p.get('profile_id')
        if pid not in seen_ids:
            seen_ids.add(pid)
            unique_profiles.append(p)
    
    # Remove near-duplicates based on buyer name
    final_profiles = []
    seen_names = {}
    for p in unique_profiles:
        name_lower = p.get('buyer_name', '').lower().strip()
        if name_lower not in seen_names:
            seen_names[name_lower] = p
            final_profiles.append(p)
        else:
            existing = seen_names[name_lower]
            existing_details = len(existing.get('buyer_details', []))
            new_details = len(p.get('buyer_details', []))
            if new_details > existing_details or (p.get('gstin') and not existing.get('gstin')):
                final_profiles.remove(existing)
                final_profiles.append(p)
                seen_names[name_lower] = p
    
    final_profiles.sort(key=lambda x: x.get('buyer_name', '').lower())
    
    if save_data(BUYER_PROFILES_JSON, final_profiles):
        removed = original_count - len(final_profiles)
        flash(f"Cleanup complete. Removed {removed} duplicate/invalid profiles.", "success")
    else:
        flash("Error during cleanup.", "error")
    
    return redirect(url_for('list_profiles'))


# ===================== FILE DOWNLOAD ROUTES =====================

@app.route('/success')
def success():
    """Success page after Excel invoice generation."""
    filename = request.args.get('filename')
    return render_template('success.html', filename=filename, 
                          download_path=f'/generated_invoices/{filename}', 
                          is_pdf=False)


@app.route('/success_pdf')
def success_pdf():
    """Success page for PDF invoice."""
    filename = request.args.get('filename')
    return render_template('success.html', filename=filename, 
                          download_path=f'/generated_invoices_pdf/{filename}', 
                          is_pdf=True)


@app.route('/generated_invoices/<filename>')
def download_file(filename):
    """Download generated Excel invoice."""
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route('/generated_invoices_pdf/<filename>')
def download_pdf_file(filename):
    """Download generated PDF invoice."""
    return send_from_directory(PDF_OUTPUT_DIR, filename, as_attachment=True)


# ===================== MAIN =====================

if __name__ == '__main__':
    print("=" * 50)
    print("Shakambhari Enterprises Invoice Generator")
    print("=" * 50)
    
    if not TEMPLATE_EXCEL_FILE:
        print("⚠️  WARNING: No template Excel file found!")
        print("   Place a template .xlsx file inside 'GST Invoices' folder.")
    else:
        print(f"✓ Template file: {os.path.basename(TEMPLATE_EXCEL_FILE)}")
    
    print(f"✓ Output folder: {OUTPUT_DIR}")
    print(f"✓ PDF folder: {PDF_OUTPUT_DIR}")
    print(f"✓ PDF conversion: {'Available' if WIN32COM_AVAILABLE else 'Not available'}")
    print("=" * 50)
    print("Starting server at http://127.0.0.1:5000")
    print("=" * 50)
    
    app.run(debug=True, port=5000)
