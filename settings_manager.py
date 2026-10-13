"""
Dynamic Settings Manager for Shakambhari Invoice Generator
=========================================================
Handles business configuration, numbering formats, default items,
AI assistant model preferences, and template options.
Persists locally and synchronizes with Google Cloud Storage for Cloud Run.
"""

import json
import os
import logging
from pathlib import Path
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)

DEFAULT_SETTINGS: Dict[str, Any] = {
    # Business Identity
    "company_name": "Shakambhari Enterprises",
    "company_subtitle": "Aluminium Utensils Manufacturing & Trading",
    "company_address": "5th Floor, Flat 26B, Om Bhawan, 54/5A, Strand Road, Jorabagan, Kolkata - 700006, WB (19)",
    "company_gstin": "19ACZPN5725A1Z8",
    "company_state": "West Bengal",
    "company_state_code": "19",
    "company_phone": "",
    "company_dispatch_address": "144/145 J.N Mukherjee Road, Ghusuri, Howrah (Near Rajesh Weight Bridge)",

    # Invoice Numbering & Formatting
    "invoice_prefix": "",
    "invoice_number_format": "{num}/2026-27",
    "financial_year": "2026-27",
    "auto_sync_ewaybill_date": True,

    # Goods & Taxation Defaults
    "default_item_description": "Aluminium Utensils",
    "default_hsn_code": "76151030",
    "default_delivery_charge": 0.0,
    "default_tax_type": "IGST",
    "default_transport_mode": "",

    # AI Multimodal Vision Assistant (Gemini)
    "gemini_primary_model": "gemini-3.8-flash",
    "gemini_backup_models": [
        "gemini-3.7-flash",
        "gemini-3-flash-preview",
        "gemini-3.1-flash-lite",
        "gemini-3.1-flash-lite-preview",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite"
    ],
    "gemini_backend_api_key": os.environ.get("GEMINI_API_KEY", ""),

    # Template & Signatures
    "master_template_filename": "invoice_template_2026_27.xlsx",
    "signatory_title": "Authorised Signatory",
    "firm_signatory_header": "For Shakambhari Enterprises"
}

DEPRECATED_OR_DISALLOWED_MODELS = {
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
    "gemini-2.0-flash-001",
    "gemini-2.0-flash-lite",
    "gemini-2.0-flash-lite-001",
    "gemini-1.5-flash",
    "gemini-1.5-flash-001",
    "gemini-1.5-flash-002",
    "gemini-1.5-flash-8b",
    "gemini-1.5-pro",
    "gemini-2.0-pro-exp",
    "gemini-flash",
    "gemini-pro"
}

ACTIVE_STANDARD_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite"
]


def sanitize_model_settings(settings: Dict[str, Any]) -> None:
    """Filter out deprecated/disallowed models and guarantee a robust fallback queue."""
    primary = settings.get('gemini_primary_model', 'gemini-3.8-flash')
    if primary in DEPRECATED_OR_DISALLOWED_MODELS or not primary:
        primary = 'gemini-3.8-flash'
    settings['gemini_primary_model'] = primary

    raw_backups = settings.get('gemini_backup_models')
    if not isinstance(raw_backups, list):
        raw_backups = []
    
    cleaned_backups = [m for m in raw_backups if m not in DEPRECATED_OR_DISALLOWED_MODELS and m != primary]
    for m in ACTIVE_STANDARD_MODELS:
        if m != primary and m not in cleaned_backups:
            cleaned_backups.append(m)
    settings['gemini_backup_models'] = cleaned_backups

SETTINGS_FILE_LOCAL = Path(__file__).parent / "app_settings.json"
GCS_SETTINGS_PATH = "config/app_settings.json"

_cached_settings: Optional[Dict[str, Any]] = None


def get_settings(storage=None) -> Dict[str, Any]:
    """Retrieve current settings, checking cache, GCS, local file, then defaults."""
    global _cached_settings
    if _cached_settings is not None:
        return dict(_cached_settings)

    settings = dict(DEFAULT_SETTINGS)

    # 1. Try reading from GCS if storage client is provided
    if storage:
        try:
            gcs_bytes = storage.download_file(GCS_SETTINGS_PATH)
            if gcs_bytes:
                gcs_data = json.loads(gcs_bytes.decode('utf-8'))
                if isinstance(gcs_data, dict):
                    settings.update(gcs_data)
                    sanitize_model_settings(settings)
                    _cached_settings = settings
                    return dict(settings)
        except Exception as e:
            logger.warning("Could not load settings from Cloud Storage: %s", e)

    # 2. Try reading from local file
    if SETTINGS_FILE_LOCAL.exists():
        try:
            with open(SETTINGS_FILE_LOCAL, 'r', encoding='utf-8') as f:
                local_data = json.load(f)
                if isinstance(local_data, dict):
                    settings.update(local_data)
        except Exception as e:
            logger.warning("Could not load local app_settings.json: %s", e)

    sanitize_model_settings(settings)
    _cached_settings = settings
    return dict(settings)


def update_settings(updates: Dict[str, Any], storage=None) -> Dict[str, Any]:
    """Update settings in memory, save locally, and sync to GCS."""
    global _cached_settings
    current = get_settings(storage)

    # Allowed keys to update
    for k, v in updates.items():
        if k in DEFAULT_SETTINGS:
            if isinstance(DEFAULT_SETTINGS[k], float):
                try:
                    current[k] = float(v)
                except (ValueError, TypeError):
                    pass
            elif isinstance(DEFAULT_SETTINGS[k], bool):
                current[k] = bool(v)
            else:
                current[k] = v

    if 'company_name' in updates and 'firm_signatory_header' not in updates:
        current['firm_signatory_header'] = f"For {current['company_name']}"

    sanitize_model_settings(current)
    _cached_settings = current

    # Save to local file
    try:
        with open(SETTINGS_FILE_LOCAL, 'w', encoding='utf-8') as f:
            json.dump(current, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logger.error("Failed to write local app_settings.json: %s", e)

    # Sync to GCS
    if storage:
        try:
            settings_json_bytes = json.dumps(current, indent=2, ensure_ascii=False).encode('utf-8')
            storage.upload_file(settings_json_bytes, GCS_SETTINGS_PATH, content_type='application/json')
            logger.info("Successfully synced app_settings.json to GCS: %s", GCS_SETTINGS_PATH)
        except Exception as e:
            logger.error("Failed to sync settings to GCS: %s", e)

    return dict(current)
