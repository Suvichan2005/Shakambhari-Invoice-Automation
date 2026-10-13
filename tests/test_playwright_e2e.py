import sys, os, time, threading, io, re, openpyxl
sys.path.insert(0, os.path.abspath('.'))
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from app import app as flask_app
from werkzeug.serving import make_server
from playwright.sync_api import sync_playwright


class ServerThread(threading.Thread):
    def __init__(self, app, host='127.0.0.1', port=5055):
        threading.Thread.__init__(self)
        self.srv = make_server(host, port, app)
        self.ctx = app.app_context()
        self.ctx.push()

    def run(self):
        self.srv.serve_forever()

    def shutdown(self):
        self.srv.shutdown()


def run_e2e_tests():
    flask_app.config['TESTING'] = True
    flask_app.config['WTF_CSRF_ENABLED'] = False
    
    server = ServerThread(flask_app, port=5055)
    server.daemon = True
    server.start()
    time.sleep(1)

    base_url = 'http://127.0.0.1:5055'
    print(f"Server started at {base_url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})  # Mobile Viewport
        page = context.new_page()

        # TEST 1: Load Homepage and verify title and seller info
        print("\n--- TEST 1: Loading Homepage on Mobile Viewport ---")
        page.goto(f"{base_url}/")
        page.wait_for_load_state('networkidle')
        title = page.title()
        print(f"Page Title: {title}")
        assert "Shakambhari Enterprises" in title or "Tax Invoice" in title or "GST" in title

        # Verify no '[COMPANY NAME]' on the page
        content = page.content()
        assert "[COMPANY NAME]" not in content, "Found '[COMPANY NAME]' placeholder on page!"
        assert "19AAAAA0000A1Z5" not in content, "Found placeholder GSTIN 19AAAAA0000A1Z5 on page!"
        print("[OK] Homepage loads clean without placeholder leaks.")

        # TEST 2: Select a buyer from dropdown and verify details
        print("\n--- TEST 2: Selecting Buyer and Form Filling ---")
        # Search buyer
        search_input = page.locator('#buyer_profile_search')
        if search_input.is_visible():
            search_input.click()
            search_input.fill('Das Metal')
            page.wait_for_timeout(300)
            buyer_item = page.locator('.buyer-dropdown-item').first
            if buyer_item.is_visible():
                buyer_item.click()
                print("[OK] Selected Das Metal from dropdown.")
        
        page.wait_for_timeout(500)
        # Check buyer card
        buyer_card = page.locator('#buyer_selected_card')
        assert buyer_card.is_visible(), "Buyer card should be visible after selection"
        card_text = buyer_card.inner_text()
        print(f"Buyer Card Text:\n{card_text}")
        assert "[COMPANY NAME]" not in card_text

        # TEST 3: Add Items, Rates, Bags, and Delivery Charge
        print("\n--- TEST 3: Filling Invoice Form Details ---")
        # Ensure Item 1 has values
        page.locator('input[name="item_quantity[]"]').first.fill('210.4')
        page.locator('input[name="item_rate[]"]').first.fill('370')
        page.locator('input[name="item_bags[]"]').first.fill('5')
        
        # Set Transport Mode
        transport_input = page.locator('#transport_mode_input')
        if transport_input.is_visible():
            transport_input.fill('Gaya Aurangabad Transport')
        
        # Set Delivery Charge
        delivery_input = page.locator('#delivery_charge')
        if delivery_input.is_visible():
            delivery_input.fill('500')

        page.wait_for_timeout(300)

        # TEST 4: Generate Bill and verify generated invoice
        print("\n--- TEST 4: Submitting Invoice Form ---")
        # Click Generate Bill
        generate_btn = page.locator('#btn_generate_bill')
        if not generate_btn.is_visible():
            generate_btn = page.locator('button[type="submit"]').first
        
        generate_btn.click()
        page.wait_for_load_state('networkidle')
        print(f"Landed on URL: {page.url}")
        assert "success" in page.url or "preview" in page.url or page.locator('.success-container').is_visible()
        print("[OK] Form submission generated bill successfully.")

        # TEST 5: Verify Generated Excel File
        print("\n--- TEST 5: Verifying Generated Invoice Excel Structure ---")
        # Find the latest generated Excel invoice
        gen_dir = 'Generated_Invoices'
        if os.path.exists(gen_dir):
            files = [os.path.join(gen_dir, f) for f in os.listdir(gen_dir) if f.endswith('.xlsx')]
            files.sort(key=os.path.getmtime, reverse=True)
            if files:
                latest_xlsx = files[0]
                print(f"Inspecting latest generated file: {latest_xlsx}")
                wb = openpyxl.load_workbook(latest_xlsx)
                ws = wb.active
                
                # Check Seller details
                seller_name = ws['A5'].value
                seller_gstin = ws['A9'].value
                bank_name = ws['A43'].value
                sig_header = ws['G45'].value

                print(f"A5 (Seller): {seller_name}")
                print(f"A9 (GSTIN): {seller_gstin}")
                print(f"A43 (Bank): {bank_name}")
                print(f"G45 (For Firm): {sig_header}")

                assert seller_name == "M/S Shakambhari Enterprises.", f"Expected Shakambhari, got {seller_name}"
                assert seller_gstin == "G.S.T No : 19ACZPN5725A1Z8", f"Expected real GSTIN, got {seller_gstin}"
                assert bank_name == "ICICI Bank", f"Expected ICICI Bank, got {bank_name}"
                assert sig_header == "For Shakambhari Enterprises", f"Expected For Shakambhari Enterprises, got {sig_header}"
                assert "[COMPANY NAME]" not in str(seller_name)
                print("[OK] Generated Excel has 100% authentic Shakambhari details!")

        # TEST 6: Test Add/Edit Buyer Modal with Aadhaar
        print("\n--- TEST 6: Testing Add Buyer Modal with Aadhaar Details ---")
        page.goto(f"{base_url}/")
        page.wait_for_load_state('networkidle')

        add_buyer_btn = page.locator('#btn_open_add_buyer')
        if add_buyer_btn.is_visible():
            add_buyer_btn.click()
            page.wait_for_timeout(300)
            
            modal = page.locator('#modal_edit_buyer')
            assert modal.is_visible(), "Add Buyer modal should be visible"

            # Fill Aadhaar buyer
            page.locator('#modal_buyer_name').fill('Ramesh Kumar Gupta')
            page.locator('#modal_buyer_details').fill(
                "Ramesh Kumar Gupta\nS/O: Mohan Lal Gupta\nMain Road, Gaya, Bihar - 823001\nState: Bihar"
            )
            page.locator('#modal_buyer_aadhaar').fill('7234 5678 9012')
            page.locator('#modal_buyer_state').fill('Bihar')
            page.locator('#modal_buyer_state_code').fill('10')

            # Save
            page.locator('#btn_modal_save_buyer').click()
            page.wait_for_timeout(500)
            print("[OK] Saved Aadhaar buyer through modal.")

            # Verify selected buyer details
            card_text = page.locator('#buyer_selected_card').inner_text()
            print(f"Selected Aadhaar Buyer Card:\n{card_text}")
            assert "Mohan Lal Gupta" not in card_text, "Father name should be stripped!"
            assert "7234 5678 9012" in card_text, "Aadhaar should be present!"
            assert "STATE- BIHAR CODE-10" in card_text, "State code 10 should be present!"
            print("[OK] Aadhaar buyer normalized cleanly in UI.")

        browser.close()
        server.shutdown()
        print("\n==========================================")
        print("ALL PLAYWRIGHT E2E TESTS PASSED 100%!")
        print("==========================================")


if __name__ == '__main__':
    run_e2e_tests()
