from playwright.sync_api import expect
import re

def test_admin_dashboard_loads(admin_page, base_url):
    admin_page.goto(f"{base_url}/admin")
    
    # Check if dashboard elements load
    expect(admin_page.locator("h1.page-title")).to_contain_text("Dashboard")
    
    # Check for KPI cards
    expect(admin_page.locator(".kpi-grid")).to_be_visible()

def test_navigation_to_spares(admin_page):
    # Click on the Spare Parts link in the sidebar
    admin_page.click("a.nav-secondary:has-text('Spare Parts')")
    
    # Verify the URL changed and the Spare Parts page loaded
    expect(admin_page).to_have_url(re.compile(r".*/admin/spares"))
    expect(admin_page.locator("h1.page-title")).to_contain_text("Spare Parts Inventory")

def test_navigation_to_cylinders(admin_page):
    admin_page.click("a.nav-secondary:has-text('Cylinders')")
    expect(admin_page).to_have_url(re.compile(r".*/admin/cylinders"))
    expect(admin_page.locator("h1.page-title")).to_contain_text("Cylinder Registry")
