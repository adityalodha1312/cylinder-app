import pytest
from playwright.sync_api import Page

@pytest.fixture
def admin_page(page: Page, base_url: str):
    """
    A fixture that navigates to the live app and logs in as the admin.
    """
    if not base_url:
        pytest.fail("You must provide --base-url to run tests against the live app. e.g. pytest --base-url https://your-app.onrender.com")

    # Go to login page
    page.goto(f"{base_url}/login")
    
    # Fill in credentials provided by the user
    page.fill('input[name="username"]', 'admin')
    page.fill('input[name="password"]', 'admin')
    page.click('button[type="submit"]')
    
    # Wait for navigation to dashboard to confirm login
    page.wait_for_url("**/admin**")
    
    yield page
