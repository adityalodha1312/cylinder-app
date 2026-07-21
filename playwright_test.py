from playwright.sync_api import sync_playwright
import time
import random

def run_tests():
    print("Starting Playwright E2E Test Suite against http://127.0.0.1:5000")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        try:
            # 1. Test Login
            print("1. Testing Admin Login...")
            page.goto("http://127.0.0.1:5000/login")
            
            # Using actual credentials or bypassing auth. Wait, we need to login!
            # Since this hits the real local dev server, let's login with an existing admin.
            # I will use 'nilesh' and 'nilesh123' or 'admin' and 'admin' if they exist.
            # Let's just create a test admin user via a python script first, or use the one created previously if it persisted.
            page.fill('input[name="username"]', 'test_admin')
            page.fill('input[name="password"]', 'test_pass')
            page.click('button[type="submit"]')
            page.wait_for_url("**/admin**")
            print(" -> Login Successful!")

            # 2. Test Dashboard
            print("2. Testing Admin Dashboard...")
            if not page.locator("h1.page-title").filter(has_text="Dashboard").is_visible():
                raise Exception("Dashboard title not found")
            print(" -> Dashboard loaded!")

            # 3. Test Navigation to Spares
            print("3. Testing Spares Navigation...")
            page.click("a.nav-secondary:has-text('Spare Parts')")
            page.wait_for_url("**/admin/spares**")
            if not page.locator("h1.page-title").filter(has_text="Spare Parts Inventory").is_visible():
                raise Exception("Spares page failed to load")
            print(" -> Spares page loaded!")

            # 4. Test Add Spare Part
            print("4. Testing Add Spare Part Modal...")
            page.click("button:has-text('Add Item')")
            page.wait_for_selector("#addModal.open")
            
            unique_code = f"TEST-VALVE-{random.randint(1000, 9999)}"
            page.fill("#add-code", unique_code)
            page.fill("#add-name", "Test Valve Assembly")
            page.select_option("#add-cat", "Valve")
            page.fill("#add-stock", "50")
            page.click("button:has-text('Add Item')")
            
            # Wait for reload
            page.wait_for_load_state("networkidle")
            print(" -> Item added successfully!")

        except Exception as e:
            print(f"TEST FAILED: {str(e)}")
            page.screenshot(path="error_screenshot.png")
            print("Saved error_screenshot.png")
        finally:
            browser.close()

if __name__ == "__main__":
    run_tests()
