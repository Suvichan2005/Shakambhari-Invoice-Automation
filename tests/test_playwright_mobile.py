import os
import sys
import time
import json
import threading
import unittest
from datetime import datetime
import pytest

from playwright.sync_api import sync_playwright

import app as local_app
from app import app as flask_app, get_ist_today_str

import socket

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


class PlaywrightServerThread(threading.Thread):
    def __init__(self, port):
        super().__init__(daemon=True)
        self.port = port
        flask_app.config['TESTING'] = True
        flask_app.config['WTF_CSRF_ENABLED'] = False

    def run(self):
        try:
            flask_app.run(host='127.0.0.1', port=self.port, debug=False, use_reloader=False)
        except Exception as e:
            print("Server thread stopped:", e)


@pytest.fixture(scope="module")
def live_server():
    port = find_free_port()
    base_url = f"http://127.0.0.1:{port}"
    server = PlaywrightServerThread(port)
    server.start()
    # Wait for server to come up
    import urllib.request
    started = False
    for _ in range(50):
        try:
            with urllib.request.urlopen(f"{base_url}/health", timeout=3) as resp:
                if resp.status == 200:
                    started = True
                    break
        except Exception:
            time.sleep(0.3)
    if not started:
        raise RuntimeError("Flask live server failed to start for Playwright.")
    yield base_url


def test_mobile_date_buttons_and_ist(live_server):
    """Test Date inputs and Today/Tomorrow shortcuts in mobile viewport."""
    today_ist = get_ist_today_str()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})
        page = context.new_page()

        page.goto(live_server)
        page.wait_for_selector("#invoice_date")

        # Initial value should be IST today
        initial_val = page.input_value("#invoice_date")
        assert initial_val == today_ist, f"Expected {today_ist}, got {initial_val}"

        # Click Today button
        page.click("button:has-text('Today')")
        today_val = page.input_value("#invoice_date")
        assert today_val == today_ist, f"Today button set {today_val}, expected {today_ist}"

        # Click Tomorrow button
        page.click("button:has-text('Tomorrow')")
        tomorrow_val = page.input_value("#invoice_date")
        # Ensure tomorrow is strictly after today
        d_today = datetime.strptime(today_ist, "%Y-%m-%d")
        d_tom = datetime.strptime(tomorrow_val, "%Y-%m-%d")
        assert (d_tom - d_today).days == 1, f"Tomorrow {tomorrow_val} should be 1 day after {today_ist}"

        browser.close()


def test_mobile_form_field_labels_sizes_and_defaults(live_server):
    """Test mobile card fields: min height >= 48px, unit defaults to Bags, no truncated labels."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})
        page = context.new_page()

        page.goto(live_server)
        page.wait_for_selector("#items_container")

        # Check existing item row
        item_rows = page.locator("#items_container tr.item-row")
        assert item_rows.count() >= 1

        first_row = item_rows.first
        # Unit dropdown should default to Bags
        unit_select = first_row.locator("select[name='item_unit[]']")
        assert unit_select.count() == 1
        assert unit_select.input_value() == "Bags"

        # Check options: Bags, Bundles, Boxes, Pieces
        options = unit_select.locator("option").all_inner_texts()
        assert "Bags" in options
        assert "Bundles" in options
        assert "Boxes" in options
        assert "Pieces" in options

        # Verify touch target heights >= 48px
        desc_input = first_row.locator("input[name='item_description[]']")
        box = desc_input.bounding_box()
        assert box and box['height'] >= 46, f"Item description height {box['height']} too small"

        bags_input = first_row.locator("input[name='item_bags[]']")
        bags_box = bags_input.bounding_box()
        assert bags_box and bags_box['height'] >= 46

        # Check labels text - clean English, no bracketed Hindi
        labels = first_row.locator(".mobile-field-label").all_inner_texts()
        assert any("Item Name" in l for l in labels)
        assert any("Bags" in l for l in labels)
        assert any("Unit" in l for l in labels)
        assert any("Weight (Kg)" in l for l in labels)
        assert any("Rate (₹)" in l for l in labels)
        assert not any("बोरी" in l for l in labels)
        assert not any("बंडल" in l for l in labels)

        # Enter weight and rate, check live preview calculation
        qty_input = first_row.locator("input[name='item_quantity[]']")
        rate_input = first_row.locator("input[name='item_rate[]']")
        qty_input.fill("84")
        rate_input.fill("380")
        qty_input.dispatch_event("input")
        rate_input.dispatch_event("input")

        # Amount should show 84 * 380 = 31920.00
        amount_text = first_row.locator(".row-amount-display").inner_text()
        assert "31920" in amount_text

        browser.close()


def test_ai_hero_card_and_modal_actions(live_server):
    """Test Hero Card and AI Modal buttons, title, and actions."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})
        page = context.new_page()

        page.goto(live_server)

        # Hero Card presence
        hero = page.locator(".ai-hero-card")
        assert hero.is_visible()
        assert "Chit & Slip Assistant" in hero.inner_text()
        assert "Take Photo" in hero.inner_text()

        # Check hero buttons
        btn_photo = page.locator(".btn-hero-photo")
        btn_upload = page.locator(".btn-hero-upload")
        btn_speak = page.locator(".btn-hero-speak")
        assert btn_photo.is_visible()
        assert btn_upload.is_visible()
        assert btn_speak.is_visible()

        # Open modal
        hero.click()
        ai_modal = page.locator("#ai_modal")
        assert ai_modal.is_visible()

        # Verify modern title and subtitle in modal
        modal_title = page.locator("#ai_modal .modal-title").inner_text()
        assert "Chit & Slip Assistant" in modal_title

        # Close modal
        page.click(".modal-close-btn")
        assert not ai_modal.is_visible()

        browser.close()


def test_sequential_queue_back_to_back_flow(live_server):
    """Test full multi-bill queue flow:
    1. Simulate AI extracting 2 bills: Bill 1 (Das Metal) and Bill 2 (Manik Store).
    2. Click 'Fill Bill 1'.
    3. Verify Bill 1 loaded into form and queue banner shown: 'Bill 1 of 2: Das Metal (Next up: Manik Store)'.
    4. Simulate landing on success.html with pending_invoices in sessionStorage.
    5. Verify prominent 'Continue to Next Bill: Manik Store (Bill 2 of 2)' button.
    6. Click continue -> lands back on index.html with Manik Store pre-filled!
    7. Verify queue banner shows 'Bill 2 of 2: Manik Store (Final bill on chit)'.
    """
    BASE_URL = live_server
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})
        page = context.new_page()

        page.goto(live_server)

        # Simulate multi-bill extraction in client
        two_bills_payload = {
            "success": True,
            "conversational_message": "Found 2 bills from paper chit.",
            "invoices": [
                {
                    "invoice_number": "101/2026-27",
                    "invoice_date": get_ist_today_str(),
                    "buyer_name": "Das Metal Liluah",
                    "transport_mode": "By Road",
                    "tax_type": "IGST",
                    "items": [
                        {"description": "Aluminium Utensils", "bags": "2", "unit": "Bags", "quantity": 84.0, "rate": 380.0}
                    ]
                },
                {
                    "invoice_number": "102/2026-27",
                    "invoice_date": get_ist_today_str(),
                    "buyer_name": "Manik Store Kolkata",
                    "transport_mode": "By Hand Rickshaw",
                    "tax_type": "CGST_SGST",
                    "items": [
                        {"description": "Aluminium Circles", "bags": "1", "unit": "Bundles", "quantity": 42.0, "rate": 385.0}
                    ]
                }
            ]
        }

        # Open modal and render bubble
        page.evaluate("""(data) => {
            openAiModal();
            renderAiAssistantThread(data);
        }""", two_bills_payload)

        # Verify both bill cards and batch button are rendered
        bill_cards = page.locator(".ai-bill-card")
        assert bill_cards.count() == 2
        batch_btn = page.locator("#btn_generate_all_batch")
        assert batch_btn.is_visible()
        assert "Generate All (2 Bills)" in batch_btn.inner_text()

        # Click Fill Bill 1
        bill_cards.first.locator(".btn-ai-fill-bill").click()

        # Modal should close, form should be active
        assert not page.locator("#ai_modal").is_visible()

        # Check form values for Bill 1
        assert page.input_value("#invoice_number") == "101/2026-27"
        assert page.locator("input[name='item_quantity[]']").first.input_value() == "84"
        assert page.locator("input[name='item_rate[]']").first.input_value() == "380"

        # Check Queue Banner
        banner = page.locator("#queue_status_banner")
        assert banner.is_visible()
        banner_text = banner.inner_text()
        assert "Bill 1 of 2" in banner_text
        assert "Das Metal Liluah" in banner_text
        assert "Next up: Manik Store Kolkata" in banner_text

        # Verify sessionStorage has pending_invoices with Manik Store
        pending_str = page.evaluate("() => sessionStorage.getItem('pending_invoices')")
        assert pending_str is not None
        pending = json.loads(pending_str)
        assert len(pending) == 1
        assert pending[0]['buyer_name'] == "Manik Store Kolkata"

        # Now navigate to success.html as Dad would after submitting Bill 1
        page.goto(f"{BASE_URL}/success?filename=Invoice_101-2026-27_Test.xlsx")

        # Success page should detect pending_invoices and display the Next Bill Queue Card
        queue_card = page.locator("#next_bill_queue_card")
        assert queue_card.is_visible()
        assert "Next Bill Ready on Your Chit" in queue_card.inner_text()
        assert "Manik Store Kolkata" in queue_card.inner_text()
        assert "Bill 2 of 2" in queue_card.inner_text()

        continue_btn = page.locator("#btn_continue_next_bill")
        assert continue_btn.is_visible()
        assert "Manik Store Kolkata" in continue_btn.inner_text()

        # Click 'Continue to Next Bill: Manik Store (Bill 2 of 2)'
        continue_btn.click()

        # Should navigate back to index.html and pre-fill Bill 2!
        page.wait_for_url(f"{BASE_URL}/")
        page.wait_for_selector("#invoice_number")

        # Verify Bill 2 is automatically loaded!
        assert page.input_value("#invoice_number") == "102/2026-27"
        assert page.locator("input[name='item_quantity[]']").first.input_value() == "42"
        assert page.locator("input[name='item_rate[]']").first.input_value() == "385"
        assert page.locator("select[name='item_unit[]']").first.input_value() == "Bundles"

        # Verify Queue Banner on Bill 2
        banner2 = page.locator("#queue_status_banner")
        assert banner2.is_visible()
        banner2_text = banner2.inner_text()
        assert "Bill 2 of 2" in banner2_text
        assert "Manik Store Kolkata" in banner2_text
        assert "Final bill on chit" in banner2_text

        # Verify pending_invoices in sessionStorage is now empty (since Bill 2 was the last)
        pending2_str = page.evaluate("() => sessionStorage.getItem('pending_invoices')")
        assert json.loads(pending2_str) == []

        # Now Dad submits Bill 2 and lands on success.html again
        page.goto(f"{BASE_URL}/success?filename=Invoice_102-2026-27_Test.xlsx")

        # Since queue is complete, #next_bill_queue_card should NOT be visible!
        assert not page.locator("#next_bill_queue_card").is_visible()

        browser.close()


def test_batch_generation_modal_submission(live_server):
    """Test 1-Click 'Generate All (2 Bills)' flow calling /api/generate_batch."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})
        page = context.new_page()

        page.goto(live_server)

        two_bills_payload = {
            "success": True,
            "conversational_message": "Found 2 bills.",
            "invoices": [
                {
                    "buyer_name": "Batch Das Metal",
                    "items": [{"description": "Aluminium Utensils", "bags": "2", "unit": "Bags", "quantity": 84.0, "rate": 380.0}],
                    "tax_type": "IGST"
                },
                {
                    "buyer_name": "Batch Rishab Metal",
                    "items": [{"description": "Aluminium Circles", "bags": "1", "unit": "Bundles", "quantity": 42.0, "rate": 385.0}],
                    "tax_type": "IGST"
                }
            ]
        }

        # Render in AI modal
        page.evaluate("""(data) => {
            openAiModal();
            renderAiAssistantThread(data);
        }""", two_bills_payload)

        # Click Generate All
        batch_btn = page.locator("#btn_generate_all_batch")
        assert batch_btn.is_visible()
        batch_btn.click()

        # Success modal should pop up
        page.wait_for_selector("#batch_success_modal", state="visible", timeout=15000)
        modal = page.locator("#batch_success_modal")
        assert "All 2 Bills Generated" in modal.inner_text()
        assert "Batch Das Metal" in modal.inner_text()
        assert "Batch Rishab Metal" in modal.inner_text()

        browser.close()
