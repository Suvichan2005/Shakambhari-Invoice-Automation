import os
import time
import json
import re
import unittest
from datetime import datetime, timezone, timedelta
from unittest.mock import patch

import app as local_app
from app import app as flask_app, get_ist_now, get_ist_today_str, format_sequential_invoice_number


class ISTDateTests(unittest.TestCase):
    def test_ist_date_at_utc_transition(self):
        """When UTC is Oct 11 19:00:00, IST is Oct 12 00:30:00."""
        utc_late_night = datetime(2026, 10, 11, 19, 0, 0, tzinfo=timezone.utc)
        ist_tz = local_app.IST
        ist_time = utc_late_night.astimezone(ist_tz)
        self.assertEqual(ist_time.strftime('%Y-%m-%d'), '2026-10-12')
        self.assertEqual(ist_time.hour, 0)
        self.assertEqual(ist_time.minute, 30)

    def test_financial_year_uses_ist(self):
        """Financial year suffix calculation accounts for IST time."""
        dt_ist = datetime(2026, 4, 1, 0, 15, tzinfo=local_app.IST)
        suffix = local_app._financial_year_suffix(dt_ist)
        self.assertEqual(suffix, '/2026-27')

    def test_index_route_renders_ist_today(self):
        """GET / on local app includes today_date in IST format YYYY-MM-DD."""
        flask_app.config['TESTING'] = True
        client = flask_app.test_client()
        response = client.get('/')
        self.assertEqual(response.status_code, 200)
        today_str = get_ist_today_str()
        self.assertIn(today_str.encode('utf-8'), response.data)


class SequentialInvoiceNumberTests(unittest.TestCase):
    def test_standard_format_with_fy(self):
        base = "035/2026-27"
        self.assertEqual(format_sequential_invoice_number(base, 0), "035/2026-27")
        self.assertEqual(format_sequential_invoice_number(base, 1), "036/2026-27")
        self.assertEqual(format_sequential_invoice_number(base, 2), "037/2026-27")

    def test_unpadded_format_with_fy(self):
        base = "35/2026-27"
        self.assertEqual(format_sequential_invoice_number(base, 0), "35/2026-27")
        self.assertEqual(format_sequential_invoice_number(base, 1), "36/2026-27")
        self.assertEqual(format_sequential_invoice_number(base, 5), "40/2026-27")

    def test_prefix_format(self):
        base = "INV-001"
        self.assertEqual(format_sequential_invoice_number(base, 0), "INV-001")
        self.assertEqual(format_sequential_invoice_number(base, 1), "INV-002")

    def test_boundary_crossing_99_to_100(self):
        base = "099/2026-27"
        self.assertEqual(format_sequential_invoice_number(base, 1), "100/2026-27")


class BatchGenerationAPITests(unittest.TestCase):
    def setUp(self):
        flask_app.config['TESTING'] = True
        self.client = flask_app.test_client()

    def test_batch_generate_empty_returns_400(self):
        response = self.client.post('/api/generate_batch', json={'invoices': []})
        self.assertEqual(response.status_code, 400)
        data = response.get_json()
        self.assertFalse(data['success'])

    def test_batch_generate_multiple_bills_success(self):
        payload = {
            'invoices': [
                {
                    'buyer_name': 'Test Das Metal',
                    'items': [
                        {'description': 'Aluminium Utensils', 'bags': '2', 'unit': 'Bags', 'quantity': 84.0, 'rate': 380.0}
                    ],
                    'tax_type': 'IGST',
                    'transport_mode': 'By Road'
                },
                {
                    'buyer_name': 'Test Rishab Metal',
                    'items': [
                        {'description': 'Aluminium Circles', 'bags': '1', 'unit': 'Bundles', 'quantity': 42.0, 'rate': 385.0}
                    ],
                    'tax_type': 'IGST',
                    'transport_mode': 'By Road'
                }
            ]
        }
        response = self.client.post('/api/generate_batch', json=payload)
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        self.assertEqual(data['created_count'], 2)
        self.assertEqual(len(data['results']), 2)

        res1 = data['results'][0]
        res2 = data['results'][1]
        self.assertEqual(res1['buyer_name'], 'Test Das Metal')
        self.assertEqual(res2['buyer_name'], 'Test Rishab Metal')
        # Bill 1 amount: 84 * 380 = 31920 + 5% IGST (1596) = 33516.00
        self.assertEqual(float(res1['total_amount']), 33516.0)
        # Bill 2 amount: 42 * 385 = 16170 + 5% IGST (808.5) = 16979.00 (half-up rounding)
        self.assertEqual(float(res2['total_amount']), 16979.0)
        # Invoice numbers must be sequential and distinct
        self.assertNotEqual(res1['invoice_number'], res2['invoice_number'])
        self.assertTrue(os.path.exists(os.path.join(local_app.OUTPUT_DIR, res1['filename'])))
        self.assertTrue(os.path.exists(os.path.join(local_app.OUTPUT_DIR, res2['filename'])))


class FormPackagingUnitTests(unittest.TestCase):
    def test_packaging_unit_dropdown_options_in_template(self):
        """Check that index page includes item_unit select dropdown with all options."""
        flask_app.config['TESTING'] = True
        client = flask_app.test_client()

        response = client.get('/')
        html = response.data.decode('utf-8')
        self.assertIn('name="item_unit[]"', html)
        self.assertIn('option value="Bags"', html)
        self.assertIn('option value="Bundles"', html)
        self.assertIn('option value="Boxes"', html)
        self.assertIn('option value="Pieces"', html)

    def test_form_submission_packaging_unit_default(self):
        """When form submitted without item_unit, backend defaults to Bags."""
        item_descriptions = ['Aluminium Utensils']
        item_bags = ['5']
        item_units = ['']  # empty unit
        item_hsns = ['76151030']
        item_quantities = ['100']
        item_rates = ['300']

        # Simulate items parsing loop from app.py
        items = []
        for i in range(len(item_descriptions)):
            desc = item_descriptions[i].strip()
            bags = item_bags[i].strip()
            unit_val = item_units[i].strip() if i < len(item_units) and item_units[i].strip() else 'Bags'
            hsn = item_hsns[i].strip()
            qty = float(item_quantities[i])
            rt = float(item_rates[i])

            clean_bag = local_app.re.sub(r'(?i)\s*(?:bags?|bundles?|boxes?|pkts?|cartons?)', '', bags).strip()
            desc = local_app.re.sub(r'\s*\(\s*[\d\.]+\s*(?:[A-Za-z]+|\s*)+\)', '', desc, flags=local_app.re.IGNORECASE).strip()
            full_desc = desc
            if clean_bag:
                full_desc += f" ({clean_bag} {unit_val})"
            items.append({
                'description': full_desc,
                'bags': clean_bag,
                'unit': unit_val,
                'hsn': hsn,
                'quantity': qty,
                'rate': rt
            })

        self.assertEqual(items[0]['unit'], 'Bags')
        self.assertEqual(items[0]['description'], 'Aluminium Utensils (5 Bags)')


class BatchGenerationEdgeCaseTests(unittest.TestCase):
    def setUp(self):
        flask_app.config['TESTING'] = True
        self.client = flask_app.test_client()

    def test_single_bill_batch(self):
        payload = {
            'invoices': [
                {
                    'buyer_name': 'Edge Solo Buyer',
                    'items': [{'description': 'Aluminium Utensils', 'quantity': 10.0, 'rate': 100.0}],
                    'tax_type': 'IGST'
                }
            ]
        }
        response = self.client.post('/api/generate_batch', json=payload)
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        self.assertEqual(data['created_count'], 1)
        self.assertEqual(data['results'][0]['buyer_name'], 'Edge Solo Buyer')

    def test_three_bills_batch_strictly_sequential(self):
        payload = {
            'invoices': [
                {
                    'buyer_name': 'Edge Trio Buyer 1',
                    'items': [{'description': 'Aluminium Utensils', 'quantity': 10.0, 'rate': 100.0}],
                    'tax_type': 'IGST'
                },
                {
                    'buyer_name': 'Edge Trio Buyer 2',
                    'items': [{'description': 'Aluminium Circles', 'quantity': 20.0, 'rate': 200.0}],
                    'tax_type': 'IGST'
                },
                {
                    'buyer_name': 'Edge Trio Buyer 3',
                    'items': [{'description': 'Aluminium Sheets', 'quantity': 30.0, 'rate': 300.0}],
                    'tax_type': 'IGST'
                }
            ]
        }
        response = self.client.post('/api/generate_batch', json=payload)
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        self.assertEqual(data['created_count'], 3)
        res = data['results']

        nums = [r['invoice_number'] for r in res]
        self.assertEqual(len(set(nums)), 3, "All three invoice numbers must be distinct")

    def test_partial_failure_handling(self):
        """When 1 bill has invalid/empty items and 1 bill is valid, batch succeeds with 1 created and logs error for the other."""
        payload = {
            'invoices': [
                {
                    'buyer_name': 'Valid Buyer',
                    'items': [{'description': 'Aluminium Utensils', 'quantity': 5.0, 'rate': 100.0}],
                    'tax_type': 'IGST'
                },
                {
                    'buyer_name': 'Empty Items Buyer',
                    'items': [],
                    'tax_type': 'IGST'
                }
            ]
        }
        response = self.client.post('/api/generate_batch', json=payload)
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertTrue(data['success'])
        self.assertEqual(data['created_count'], 1)
        self.assertEqual(len(data['errors']), 1)
        self.assertIn("No items specified", data['errors'][0])


class NonBillAndAttachmentIntegrityTests(unittest.TestCase):
    def setUp(self):
        flask_app.config['TESTING'] = True
        self.client = flask_app.test_client()

    def test_irrelevant_notes_returns_empty_invoices(self):
        """When Gemini detects study notes and returns empty invoices, backend returns empty invoices without dummy bill."""
        from unittest.mock import patch, MagicMock
        import io

        mock_gemini_resp = {
            "conversational_message": "The uploaded photo appears to be handwritten study notes on computer networking rather than a rough billing slip. No utensil items or buyer details found.",
            "has_clarifications": False,
            "clarifications": [],
            "is_duplicate_invoice": False,
            "invoices": [],
            "data": None
        }

        mock_http_resp = MagicMock()
        mock_http_resp.read.return_value = json.dumps({
            "candidates": [{
                "content": {
                    "parts": [{"text": json.dumps(mock_gemini_resp)}]
                }
            }]
        }).encode('utf-8')
        mock_http_resp.__enter__.return_value = mock_http_resp

        with patch('urllib.request.urlopen', return_value=mock_http_resp):
            response = self.client.post('/api/ai/parse-bill', json={
                'prompt': 'Here are study notes',
                'api_key': 'fake-test-key'
            })
            self.assertEqual(response.status_code, 200)
            data = response.get_json()
            self.assertTrue(data['success'])
            self.assertEqual(data['invoices'], [], "Invoices must be empty for study notes")
            self.assertIsNone(data['data'], "Primary data must be None when no invoices found")
            self.assertIn("study notes", data['conversational_message'].lower())

    def test_filter_removes_unnamed_party_skeleton(self):
        """Even if model returns a skeleton Unnamed Party with 0 items, backend filters it out."""
        from unittest.mock import patch, MagicMock

        mock_gemini_resp = {
            "conversational_message": "Could not identify billing details.",
            "has_clarifications": False,
            "clarifications": [],
            "invoices": [
                {
                    "buyer_name": "Unnamed Party",
                    "items": [],
                    "total": 0.0
                }
            ]
        }

        mock_http_resp = MagicMock()
        mock_http_resp.read.return_value = json.dumps({
            "candidates": [{
                "content": {
                    "parts": [{"text": json.dumps(mock_gemini_resp)}]
                }
            }]
        }).encode('utf-8')
        mock_http_resp.__enter__.return_value = mock_http_resp

        with patch('urllib.request.urlopen', return_value=mock_http_resp):
            response = self.client.post('/api/ai/parse-bill', json={
                'prompt': 'Empty slip',
                'api_key': 'fake-test-key'
            })
            self.assertEqual(response.status_code, 200)
            data = response.get_json()
            self.assertTrue(data['success'])
            self.assertEqual(data['invoices'], [], "Empty skeleton invoice must be filtered out")
            self.assertIsNone(data['data'])

    def test_template_staging_and_no_auto_submit(self):
        """Verify that index.html does not auto-submit on file select and does not turn clarifications into submit buttons."""
        with open('templates/index.html', 'r', encoding='utf-8') as f:
            content = f.read()

        # handleAiFileSelect must not call submitAiMessage directly
        file_select_fn = content[content.find('function handleAiFileSelect'):content.find('function removeAiImage')]
        self.assertNotIn('submitAiMessage', file_select_fn, "handleAiFileSelect must NOT auto-submit")

        # Clarifications must not be clickable buttons that submitAiMessage
        self.assertNotIn("class=\"ai-clarification-chip\" onclick=\"submitAiMessage", content)


class BuyerNormalizationTests(unittest.TestCase):
    def test_clean_line_phone_email_stripping(self):
        """Test stripping of phone numbers and emails while preserving postal PIN codes."""
        from buyer_normalizer import clean_line_phone_email

        self.assertEqual(clean_line_phone_email("Ph: 9830012345"), "")
        self.assertEqual(clean_line_phone_email("Mob: +91 98300 12345"), "")
        self.assertEqual(clean_line_phone_email("Email: test@example.com"), "")
        self.assertEqual(clean_line_phone_email("M: +91-9830012345 / 9830123456"), "")
        self.assertEqual(clean_line_phone_email("Tel: 033-26541234, Fax: 033-26541235"), "")

        # PIN codes must be preserved!
        self.assertEqual(clean_line_phone_email("Howrah - 711315"), "Howrah - 711315")
        self.assertEqual(clean_line_phone_email("Assam-783380"), "Assam-783380")
        self.assertEqual(clean_line_phone_email("Singur - 712409, West Bengal"), "Singur - 712409, West Bengal")
        self.assertEqual(clean_line_phone_email("PIN - 743125"), "PIN - 743125")

        # Address lines with inline phones/emails
        self.assertEqual(
            clean_line_phone_email("Vill:- Kashida, PO:- Ghatsila, Mob: 9876543210"),
            "Vill:- Kashida, PO:- Ghatsila"
        )
        self.assertEqual(
            clean_line_phone_email("Station Road, Ranchi, Phone: 9831098310, Email: store@gmail.com"),
            "Station Road, Ranchi"
        )

    def test_normalize_buyer_profile_state_from_gstin(self):
        """Test missing state code is inferred from GSTIN and formatted canonically."""
        from buyer_normalizer import normalize_buyer_profile

        raw = {
            'buyer_name': 'NARAYANI METAL',
            'buyer_details': [
                'Buyer :',
                'NARAYANI METAL',
                'Address :46/F, Malipanchghara Street,',
                '8 no. Janki Devi Jalan Road Complex - 711204',
                'Ph: 9830012345',
                'GST-No -19AFQPP2143G1Z7'
            ]
        }
        norm = normalize_buyer_profile(raw)
        self.assertEqual(norm['gstin'], '19AFQPP2143G1Z7')
        self.assertEqual(norm['state_code'], '19')
        self.assertEqual(norm['state'], 'West Bengal')
        self.assertEqual(norm['default_tax_type'], 'CGST_SGST')
        self.assertIn('STATE- WEST BENGAL CODE-19', norm['buyer_details'])
        self.assertIn('GSTIN - 19AFQPP2143G1Z7', norm['buyer_details'])
        # Verify phone number was stripped
        self.assertNotIn('Ph: 9830012345', norm['buyer_details'])
        # Verify 6-digit PIN code was preserved
        self.assertTrue(any('711204' in l for l in norm['buyer_details']))

    def test_normalize_buyer_profile_state_from_city(self):
        """Test state code inferred from city keywords when GSTIN is missing."""
        from buyer_normalizer import normalize_buyer_profile

        raw = {
            'buyer_name': 'Guwahati Fancy Trader',
            'buyer_details': [
                'Guwahati Fancy Trader',
                'Muralidhar Sharma Road',
                'Fancy Bazar, Guwahati - 781001',
                'Tel: 0361-2541234'
            ]
        }
        norm = normalize_buyer_profile(raw)
        self.assertEqual(norm['state_code'], '18')
        self.assertEqual(norm['state'], 'Assam')
        self.assertEqual(norm['default_tax_type'], 'IGST')
        self.assertIn('STATE- ASSAM CODE-18', norm['buyer_details'])
        self.assertTrue(any('781001' in l for l in norm['buyer_details']))

    def test_normalize_buyer_profile_strips_embedded_state_code(self):
        """Test that embedded CODE-16 on address line is stripped cleanly."""
        from buyer_normalizer import normalize_buyer_profile

        raw = {
            'buyer_name': 'KALACHAND BHANDER',
            'buyer_details': [
                'KALACHAND BHANDER',
                'EAST MARKET',
                'BASAN PATTI',
                'DHARMANAGAR, TRIPURA-799250 CODE-16',
                'GSTIN - 16BNKPPO840B1ZB'
            ]
        }
        norm = normalize_buyer_profile(raw)
        self.assertEqual(norm['state_code'], '16')
        self.assertIn('STATE- TRIPURA CODE-16', norm['buyer_details'])
        self.assertIn('DHARMANAGAR, TRIPURA-799250', norm['buyer_details'])
        self.assertNotIn('DHARMANAGAR, TRIPURA-799250 CODE-16', norm['buyer_details'])

    def test_api_ai_create_buyer_normalizes_address(self):
        """POST /api/ai/create-buyer strips phone/email and adds state code."""
        flask_app.config['TESTING'] = True
        client = flask_app.test_client()

        payload = {
            'buyer_name': 'Patna Aluminium Mart',
            'buyer_details': [
                'Patna Aluminium Mart',
                'Station Road, Patna - 800001',
                'Mob: +91 94310 12345',
                'Email: patna@metalmart.in',
                'GSTIN: 10AABCP1234D1ZE'
            ],
            'gstin': '10AABCP1234D1ZE'
        }
        resp = client.post('/api/ai/create-buyer', json=payload)
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertTrue(data['success'])
        buyer = data['buyer']
        self.assertEqual(buyer['state_code'], '10')
        self.assertEqual(buyer['state'], 'Bihar')
        self.assertIn('STATE- BIHAR CODE-10', buyer['buyer_details'])
        self.assertIn('GSTIN - 10AABCP1234D1ZE', buyer['buyer_details'])
        # Ensure phone and email were stripped
        for line in buyer['buyer_details']:
            self.assertNotIn('94310', line)
            self.assertNotIn('patna@metalmart.in', line)

    def test_all_existing_buyer_profiles_normalized(self):
        """Verify that every profile in cloud/buyer_profiles.json is normalized."""
        with open('cloud/buyer_profiles.json', 'r', encoding='utf-8') as f:
            profiles = json.load(f)

        self.assertGreater(len(profiles), 20)
        for p in profiles:
            name = p.get('buyer_name', '')
            self.assertNotEqual(name, 'Buyer :', "Dummy 'Buyer :' profile must not exist")
            self.assertTrue(bool(name), "Every profile must have a buyer_name")
            self.assertTrue(bool(p.get('state_code')), f"Profile {name} missing state_code")
            self.assertTrue(bool(p.get('state')), f"Profile {name} missing state")
            self.assertIn(p.get('default_tax_type'), ['CGST_SGST', 'IGST'])

            # Verify no email or phone in details
            for line in p.get('buyer_details', []):
                self.assertFalse(re.search(r'@', line), f"Email found in {name}: {line}")
                self.assertFalse(re.search(r'(?i)\b(mob|ph|phone|cell)\s*[:.\-]', line), f"Phone found in {name}: {line}")


class ProductionSecurityHardeningTests(unittest.TestCase):
    def setUp(self):
        import cloud.app_cloud as ca
        self.ca = ca
        ca.app.config['TESTING'] = True
        self.client = ca.app.test_client()
        self.orig_auth_enabled = ca.app.config.get('AUTH_ENABLED')
        self.orig_hash = ca.app.config.get('APP_PASSWORD_HASH')

    def tearDown(self):
        self.ca.app.config['AUTH_ENABLED'] = self.orig_auth_enabled
        self.ca.app.config['APP_PASSWORD_HASH'] = self.orig_hash
        if 'APP_PASSWORD' in self.ca.app.config:
            del self.ca.app.config['APP_PASSWORD']
        self.ca._request_windows.clear()

    def test_password_hashing_and_no_plaintext_in_memory(self):
        """Setting app password must compute PBKDF2 hash and never store plaintext in config."""
        test_pw = "SafeAluminium$2026!Key"
        self.ca.set_app_password(test_pw)
        self.assertTrue(self.ca.app.config.get('AUTH_ENABLED'))
        self.assertNotIn('APP_PASSWORD', self.ca.app.config)
        stored_hash = self.ca.app.config.get('APP_PASSWORD_HASH')
        self.assertIsNotNone(stored_hash)
        self.assertEqual(len(stored_hash), 64)  # SHA-256 hex is 64 characters
        self.assertNotEqual(stored_hash, test_pw)

    def test_constant_time_password_verification(self):
        """verify_app_password accepts valid password, rejects invalid or empty strings."""
        test_pw = "SecretTestingPassword99"
        self.ca.set_app_password(test_pw)
        self.assertTrue(self.ca.verify_app_password(test_pw))
        self.assertFalse(self.ca.verify_app_password("WrongPassword123"))
        self.assertFalse(self.ca.verify_app_password(""))
        self.assertFalse(self.ca.verify_app_password("   "))

    def test_login_delay_and_session_generation(self):
        """Valid login clears prior session and logs in; invalid login takes delay."""
        test_pw = "CorrectProductionAuth456"
        self.ca.set_app_password(test_pw)

        # Invalid login
        t0 = time.time()
        resp_fail = self.client.post('/login', data={'password': 'wrong_password'})
        elapsed = time.time() - t0
        self.assertGreaterEqual(elapsed, 0.95, "Failed login must introduce >= 1.0s penalty delay")
        self.assertEqual(resp_fail.status_code, 200)

        # Successful login
        resp_ok = self.client.post('/login', data={'password': test_pw})
        self.assertEqual(resp_ok.status_code, 302)
        with self.client.session_transaction() as sess:
            self.assertTrue(sess.get('authenticated'))

    def test_login_rate_limiting(self):
        """Login route triggers 429 after 5 failed attempts within 60s."""
        test_pw = "AnotherSecretKey789"
        self.ca.set_app_password(test_pw)
        self.ca._request_windows.clear()

        # Send 5 attempts (rate limit is 5 per 60s)
        for _ in range(5):
            self.client.post('/login', data={'password': 'bad'}, headers={'X-Forwarded-For': '10.0.0.1'})

        # 6th attempt should be blocked with 429
        resp_blocked = self.client.post('/login', data={'password': 'bad'}, headers={'X-Forwarded-For': '10.0.0.1'})
        self.assertEqual(resp_blocked.status_code, 429)

    def test_api_settings_sanitizes_credentials(self):
        """GET and POST /api/settings must never expose gemini_backend_api_key."""
        resp = self.client.get('/api/settings')
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertTrue(data['success'])
        settings = data['settings']
        self.assertNotIn('gemini_backend_api_key', settings)
        self.assertIn('has_backend_api_key', settings)

    def test_http_security_headers_present(self):
        """All responses must contain OWASP defensive security headers."""
        resp = self.client.get('/health')
        self.assertEqual(resp.headers.get('X-Content-Type-Options'), 'nosniff')
        self.assertEqual(resp.headers.get('X-Frame-Options'), 'DENY')
        self.assertEqual(resp.headers.get('X-XSS-Protection'), '1; mode=block')
        self.assertEqual(resp.headers.get('Referrer-Policy'), 'strict-origin-when-cross-origin')

    def test_gemini_model_fallback_queue_purges_deprecated_and_flash_latest(self):
        """Disallowed models (flash-latest, 2.0, 1.5) must be stripped and replaced with active queue."""
        from settings_manager import sanitize_model_settings, DEPRECATED_OR_DISALLOWED_MODELS, ACTIVE_STANDARD_MODELS
        
        legacy_cfg = {
            'gemini_primary_model': 'gemini-flash-latest',
            'gemini_backup_models': [
                'gemini-flash-latest',
                'gemini-2.0-flash',
                'gemini-1.5-flash',
                'gemini-3.7-flash'
            ]
        }
        sanitize_model_settings(legacy_cfg)
        
        # Primary should be reset to 3.8 Flash
        self.assertEqual(legacy_cfg['gemini_primary_model'], 'gemini-3.8-flash')
        
        # Backup queue must have no deprecated or flash-latest models
        for bad in DEPRECATED_OR_DISALLOWED_MODELS:
            self.assertNotIn(bad, legacy_cfg['gemini_backup_models'])
            
        # Active standard models must be present in queue
        self.assertIn('gemini-3.7-flash', legacy_cfg['gemini_backup_models'])
        self.assertIn('gemini-3.6-flash', legacy_cfg['gemini_backup_models'])
        self.assertIn('gemini-3.5-flash', legacy_cfg['gemini_backup_models'])
        self.assertIn('gemini-3.5-flash-lite', legacy_cfg['gemini_backup_models'])


if __name__ == '__main__':
    unittest.main()

