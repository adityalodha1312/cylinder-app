import re
from playwright.sync_api import Page, expect

def test_admin_login(page: Page, base_url: str):
    page.goto(f"{base_url}/login")
    
    expect(page).to_have_title(re.compile("Login"))
    
    page.fill('input[name="username"]', 'admin')
    page.fill('input[name="password"]', 'admin')
    page.click('button[type="submit"]')
    
    expect(page).to_have_url(re.compile(r".*/admin.*"))
    expect(page.locator("h1.page-title")).to_contain_text("Dashboard")

def test_invalid_login(page: Page, base_url: str):
    page.goto(f"{base_url}/login")
    page.fill('input[name="username"]', 'wrong_user')
    page.fill('input[name="password"]', 'wrong_pass')
    page.click('button[type="submit"]')
    
    expect(page.locator(".error-msg")).to_be_visible()
    expect(page.locator(".error-msg")).to_contain_text("Invalid username or password")
