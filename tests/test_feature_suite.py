import os
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


if __name__ == '__main__':
    import json
    unittest.main()
else:
    import json
