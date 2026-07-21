from playwright.sync_api import expect
import time
import random
import re

def test_add_spare_part(admin_page):
    admin_page.goto(f"{admin_page.url.split('/admin')[0]}/admin/spares")
    admin_page.click("button:has-text('Add Item')")
    expect(admin_page.locator("#addModal")).to_have_class(re.compile(r".*open.*"))
    
    unique_code = f"TEST-VALVE-{random.randint(1000, 9999)}"
    admin_page.fill("#add-code", unique_code)
    admin_page.fill("#add-name", "Test Valve Assembly")
    admin_page.select_option("#add-cat", "Valve")
    admin_page.fill("#add-stock", "50")
    
    admin_page.click("button:has-text('Add Item')")
    admin_page.wait_for_load_state("networkidle")
    
    admin_page.fill("#spares-search", unique_code)
    expect(admin_page.locator("table.reg-table")).to_contain_text(unique_code)

def test_cylinder_search(admin_page):
    admin_page.goto(f"{admin_page.url.split('/admin')[0]}/admin/cylinders")
    admin_page.fill("input#search_input", "XYZ999NONEXISTENT")
    admin_page.click("button.search-btn")
    expect(admin_page.locator("table.reg-table")).to_contain_text("No cylinders found")
