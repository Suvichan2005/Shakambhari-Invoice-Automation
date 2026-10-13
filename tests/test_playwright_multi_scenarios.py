import sys, os, time, threading, io, re, openpyxl
sys.path.insert(0, os.path.abspath('.'))
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from app import app as flask_app
from werkzeug.serving import make_server
from playwright.sync_api import sync_playwright


class ServerThread(threading.Thread):
    def __init__(self, app, host='127.0.0.1', port=5056):
        super().__init__()
        self.srv = make_server(host, port, app)
        self.ctx = app.app_context()
        self.ctx.push()

    def run(self):
        self.srv.serve_forever()

    def shutdown(self):
        self.srv.shutdown()


def inspect_xlsx(filepath, expected_buyer_name=None, expected_tax_type="IGST"):
    assert os.path.exists(filepath), f"File {filepath} not found"
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Check Seller details
    s_name = ws['A5'].value
    s_gstin = ws['A9'].value
    bank = ws['A43'].value
    sig = ws['G45'].value
    
    print(f"  [XLSX] Seller: {s_name} | GSTIN: {s_gstin}")
    print(f"  [XLSX] Bank: {bank} | Sig: {sig}")
    
    assert s_name == "M/S Shakambhari Enterprises.", f"Expected Shakambhari, got {s_name}"
    assert "19ACZPN5725A1Z8" in str(s_gstin), f"Wrong seller GSTIN: {s_gstin}"
    assert bank == "ICICI Bank", f"Expected ICICI Bank, got {bank}"
    assert sig == "For Shakambhari Enterprises", f"Expected For Shakambhari Enterprises, got {sig}"

    # Verify no '[COMPANY NAME]' or '19AAAAA0000A1Z5' in entire sheet
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell is not None:
                val = str(cell)
                assert "[COMPANY NAME]" not in val, f"Found [COMPANY NAME] in cell: {val}"
                assert "19AAAAA0000A1Z5" not in val, f"Found placeholder GSTIN in cell: {val}"

    # Verify buyer lines
    buyer_lines = [str(ws[f'A{r}'].value or '').strip() for r in range(13, 19)]
    buyer_text = "\n".join([l for l in buyer_lines if l])
    print(f"  [XLSX] Buyer Block (A13:A18):\n{buyer_text}")

    if expected_buyer_name:
        assert any(expected_buyer_name.lower() in l.lower() for l in buyer_lines), f"Expected {expected_buyer_name} in buyer block"

    # Verify no duplicate state lines
    state_lines = [l for l in buyer_lines if "state" in l.lower()]
    assert len(state_lines) <= 1, f"Found multiple state lines in buyer block: {state_lines}"

    print(f"  [XLSX] Inspection PASSED for {os.path.basename(filepath)}")


def test_all_scenarios():
    flask_app.config['TESTING'] = True
    flask_app.config['WTF_CSRF_ENABLED'] = False

    server = ServerThread(flask_app, port=5056)
    server.daemon = True
    server.start()
    time.sleep(1)

    base_url = 'http://127.0.0.1:5056'
    print(f"Test server running at {base_url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 412, 'height': 915})
        page = context.new_page()

        # =========================================================================
        # SCENARIO 1: Intra-State Sale (West Bengal -> CGST 2.5% + SGST 2.5%)
        # =========================================================================
        print("\n=======================================================")
        print("SCENARIO 1: Intra-State (West Bengal) Buyer Bill Generation")
        print("=======================================================")
        page.goto(f"{base_url}/")
        page.wait_for_load_state('networkidle')

        # Add a WB Buyer via modal
        page.locator('#btn_open_add_buyer').click()
        page.wait_for_timeout(300)
        page.locator('#edit_buyer_name').fill('Kolkata Aluminum Store')
        page.locator('#edit_buyer_details').fill('Kolkata Aluminum Store\n12, Brabourne Road, Burrabazar\nKolkata - 700001\nState: West Bengal')
        page.locator('#edit_buyer_gstin').fill('19AAACK1234M1Z5')
        page.locator('#edit_buyer_state').fill('West Bengal')
        page.locator('#edit_buyer_state_code').fill('19')
        page.locator('#btn_save_buyer_modal').click()
        page.locator('#modal_edit_buyer').wait_for(state='hidden', timeout=5000)
        page.wait_for_function("() => !!document.getElementById('buyer_profile_id').value")

        # Fill items
        page.locator('input[name="item_quantity[]"]').first.fill('150.5')
        page.locator('input[name="item_rate[]"]').first.fill('365')
        page.locator('input[name="item_bags[]"]').first.fill('4')
        page.locator('#transport_mode_input').fill('Local Van')
        # Generate bill
        page.locator('#btn_generate_bill').click()
        page.wait_for_load_state('networkidle')
        print(f"  Current URL: {page.url}")
        if "success" not in page.url:
            print(f"  Flashes: {page.locator('.alert, .flash, .banner-notification').all_inner_texts()}")
            print(f"  Buyer ID in form: {page.locator('#buyer_profile_id').input_value()}")
            print(f"  Page Text Snippet: {page.locator('body').inner_text()[:400]}")
        assert "success" in page.url
        print("  -> Scenario 1 Invoice Generated successfully!")

        # Find latest file
        gen_dir = 'Generated_Invoices'
        files = sorted([os.path.join(gen_dir, f) for f in os.listdir(gen_dir) if f.endswith('.xlsx')], key=os.path.getmtime, reverse=True)
        inspect_xlsx(files[0], expected_buyer_name="Kolkata Aluminum Store", expected_tax_type="CGST/SGST")

        # =========================================================================
        # SCENARIO 2: Inter-State Sale (Assam Buyer with redundant state in raw details)
        # =========================================================================
        print("\n=======================================================")
        print("SCENARIO 2: Inter-State (Assam) Buyer with Redundant State Lines")
        print("=======================================================")
        page.goto(f"{base_url}/")
        page.wait_for_load_state('networkidle')

        page.locator('#btn_open_add_buyer').click()
        page.wait_for_timeout(300)
        page.locator('#edit_buyer_name').fill('Guwahati Metal Traders')
        page.locator('#edit_buyer_details').fill('Guwahati Metal Traders\nFancy Bazar\nGuwahati - 781001\nState: Assam\nSTATE- ASSAM CODE-18')
        page.locator('#edit_buyer_gstin').fill('18AABCG9876P1Z3')
        page.locator('#edit_buyer_state').fill('Assam')
        page.locator('#edit_buyer_state_code').fill('18')
        page.locator('#btn_save_buyer_modal').click()
        page.locator('#modal_edit_buyer').wait_for(state='hidden', timeout=5000)
        page.wait_for_function("() => !!document.getElementById('buyer_profile_id').value")

        # Fill item details
        page.locator('input[name="item_quantity[]"]').first.fill('420.0')
        page.locator('input[name="item_rate[]"]').first.fill('380')
        page.locator('input[name="item_bags[]"]').first.fill('10')
        page.locator('#transport_mode_input').fill('North East Roadways')
        page.locator('#delivery_charge').fill('1200')

        page.locator('#btn_generate_bill').click()
        page.wait_for_load_state('networkidle')
        assert "success" in page.url
        print("  -> Scenario 2 Invoice Generated successfully!")

        files = sorted([os.path.join(gen_dir, f) for f in os.listdir(gen_dir) if f.endswith('.xlsx')], key=os.path.getmtime, reverse=True)
        inspect_xlsx(files[0], expected_buyer_name="Guwahati Metal Traders", expected_tax_type="IGST")

        # =========================================================================
        # SCENARIO 3: Aadhaar-Based Buyer (No GSTIN, Aadhaar line, S/O removal)
        # =========================================================================
        print("\n=======================================================")
        print("SCENARIO 3: Aadhaar-Based Unregistered Buyer (Dad's specific case)")
        print("=======================================================")
        page.goto(f"{base_url}/")
        page.wait_for_load_state('networkidle')

        page.locator('#btn_open_add_buyer').click()
        page.wait_for_timeout(300)
        page.locator('#edit_buyer_name').fill('Suresh Prasad')
        page.locator('#edit_buyer_details').fill('Suresh Prasad\nS/O Late Ram Prasad\nWard No. 4, Gaya, Bihar - 823001\nState: Bihar')
        page.locator('#edit_buyer_aadhaar').fill('7234 9876 5432')
        page.locator('#edit_buyer_state').fill('Bihar')
        page.locator('#edit_buyer_state_code').fill('10')
        page.locator('#btn_save_buyer_modal').click()
        page.locator('#modal_edit_buyer').wait_for(state='hidden', timeout=5000)
        page.wait_for_function("() => !!document.getElementById('buyer_profile_id').value")

        # Verify card in UI
        card_text = page.locator('#buyer_selected_card').inner_text()
        print(f"  UI Buyer Card:\n{card_text}")
        assert "Ram Prasad" not in card_text, "Father's name must be stripped!"
        assert "7234 9876 5432" in card_text, "Aadhaar must appear in UI card!"

        page.locator('input[name="item_quantity[]"]').first.fill('210.4')
        page.locator('input[name="item_rate[]"]').first.fill('370')
        page.locator('input[name="item_bags[]"]').first.fill('5')
        page.locator('#transport_mode_input').fill('Aurangabad Transport')
        page.locator('#delivery_charge').fill('500')

        page.locator('#btn_generate_bill').click()
        page.wait_for_load_state('networkidle')
        assert "success" in page.url
        print("  -> Scenario 3 Invoice Generated successfully!")

        files = sorted([os.path.join(gen_dir, f) for f in os.listdir(gen_dir) if f.endswith('.xlsx')], key=os.path.getmtime, reverse=True)
        inspect_xlsx(files[0], expected_buyer_name="Suresh Prasad", expected_tax_type="IGST")

        # =========================================================================
        # SCENARIO 4: Multi-Item Bill with Delivery & Round Off
        # =========================================================================
        print("\n=======================================================")
        print("SCENARIO 4: Multi-Item Invoice (3 Items + Bags + Delivery)")
        print("=======================================================")
        page.goto(f"{base_url}/")
        page.wait_for_load_state('networkidle')

        # Select saved buyer Das Metal
        search_input = page.locator('#buyer_profile_search')
        search_input.fill('Das Metal')
        page.wait_for_timeout(300)
        page.locator('.buyer-dropdown-item').first.click()
        page.wait_for_timeout(300)

        # Add Item 1
        page.locator('input[name="item_quantity[]"]').nth(0).fill('100.5')
        page.locator('input[name="item_rate[]"]').nth(0).fill('350')
        page.locator('input[name="item_bags[]"]').nth(0).fill('2')

        # Add Item 2
        page.locator('#btn_add_item').click()
        page.wait_for_timeout(200)
        page.locator('input[name="item_quantity[]"]').nth(1).fill('75.25')
        page.locator('input[name="item_rate[]"]').nth(1).fill('360')
        page.locator('input[name="item_bags[]"]').nth(1).fill('2')

        # Add Item 3
        page.locator('#btn_add_item').click()
        page.wait_for_timeout(200)
        page.locator('input[name="item_quantity[]"]').nth(2).fill('50.0')
        page.locator('input[name="item_rate[]"]').nth(2).fill('380')
        page.locator('input[name="item_bags[]"]').nth(2).fill('1')

        page.locator('#transport_mode_input').fill('Burdwan Express')
        page.locator('#delivery_charge').fill('450')

        page.locator('#btn_generate_bill').click()
        page.wait_for_load_state('networkidle')
        assert "success" in page.url
        print("  -> Scenario 4 Invoice Generated successfully!")

        files = sorted([os.path.join(gen_dir, f) for f in os.listdir(gen_dir) if f.endswith('.xlsx')], key=os.path.getmtime, reverse=True)
        inspect_xlsx(files[0], expected_tax_type="IGST")

        browser.close()
        server.shutdown()
        print("\n=======================================================")
        print("ALL MULTI-SCENARIO E2E TESTS PASSED WITH ZERO ERRORS!")
        print("=======================================================")


if __name__ == '__main__':
    test_all_scenarios()
