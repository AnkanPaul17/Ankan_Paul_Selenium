import os
import time

import pytest
from openpyxl import load_workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoAlertPresentException,
    TimeoutException
)


# ---------------------------------------------------------
# PATHS
# ---------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(
    BASE_DIR,
    "test_data",
    "testdata.xlsx"
)

SCREENSHOT_DIR = os.path.join(
    BASE_DIR,
    "screenshots"
)

os.makedirs(SCREENSHOT_DIR, exist_ok=True)


# ---------------------------------------------------------
# READ TEST DATA FROM EXCEL
# ---------------------------------------------------------

def read_test_data():

    workbook = load_workbook(EXCEL_FILE)

    sheet = workbook.active

    email = sheet["A2"].value
    password = sheet["B2"].value
    product = sheet["C2"].value
    quantity = int(sheet["D2"].value)

    return email, password, product, quantity


# ---------------------------------------------------------
# SCREENSHOT FUNCTION
# ---------------------------------------------------------

def take_screenshot(driver, name):

    file_path = os.path.join(
        SCREENSHOT_DIR,
        name
    )

    driver.save_screenshot(file_path)

    print(f"Screenshot saved: {file_path}")


# ---------------------------------------------------------
# ALERT HANDLER
# ---------------------------------------------------------

def handle_alert(driver):

    try:

        alert = driver.switch_to.alert

        print("Alert found:")
        print(alert.text)

        alert.accept()

        print("Alert accepted.")

    except NoAlertPresentException:

        print("No alert present.")


# ---------------------------------------------------------
# SETUP
# ---------------------------------------------------------

@pytest.fixture
def driver():

    options = webdriver.ChromeOptions()

    options.add_argument("--start-maximized")

    browser = webdriver.Chrome(
        options=options
    )

    browser.implicitly_wait(5)

    yield browser

    browser.quit()


# ---------------------------------------------------------
# MAIN TEST
# ---------------------------------------------------------

def test_ecommerce_purchase(driver):

    wait = WebDriverWait(driver, 15)

    # -----------------------------------------------------
    # READ DATA FROM EXCEL
    # -----------------------------------------------------

    email, password, product, quantity = read_test_data()

    print("\nTest Data:")
    print("Email:", email)
    print("Product:", product)
    print("Quantity:", quantity)

    # -----------------------------------------------------
    # 1. LAUNCH WEBSITE
    # -----------------------------------------------------

    print("\n1. Launching website...")

    driver.get(
        "https://tutorialsninja.com/demo/"
    )

    take_screenshot(
        driver,
        "01_homepage.png"
    )

    # -----------------------------------------------------
    # 2. LOGIN
    # -----------------------------------------------------

    print("\n2. Logging into application...")

    # Click My Account
    my_account = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//a[contains(., 'My Account')]"
            )
        )
    )

    my_account.click()

    # Click Login
    login_link = wait.until(
        EC.element_to_be_clickable(
            (
                By.LINK_TEXT,
                "Login"
            )
        )
    )

    login_link.click()

    # Enter email
    email_field = wait.until(
        EC.visibility_of_element_located(
            (
                By.ID,
                "input-email"
            )
        )
    )

    email_field.clear()
    email_field.send_keys(email)

    # Enter password
    password_field = driver.find_element(
        By.ID,
        "input-password"
    )

    password_field.clear()
    password_field.send_keys(password)

    # Click Login
    login_button = driver.find_element(
        By.XPATH,
        "//input[@value='Login']"
    )

    login_button.click()

    # Check login succeeded
    wait.until(
        EC.url_contains("account/account")
    )

    print("Login successful.")

    take_screenshot(
        driver,
        "02_login_success.png"
    )

    # -----------------------------------------------------
    # 3. SEARCH PRODUCT
    # -----------------------------------------------------

    print("\n3. Searching for product:", product)

    search_box = wait.until(
        EC.visibility_of_element_located(
            (
                By.NAME,
                "search"
            )
        )
    )

    search_box.clear()

    search_box.send_keys(product)

    search_box.send_keys(Keys.RETURN)

    # Wait for search result
    wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//div[contains(@class,'product-thumb')]"
            )
        )
    )

    print("Product search completed.")

    take_screenshot(
        driver,
        "03_product_search.png"
    )

    # -----------------------------------------------------
    # 4. ADD PRODUCT TO CART
    # -----------------------------------------------------

    print("\n4. Adding product to cart...")

    # Find product card containing our product
    product_card = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                f"//div[contains(@class,'product-thumb')][.//a[contains(normalize-space(), '{product}')]]"
            )
        )
    )

    # Click Add to Cart inside product card
    add_to_cart_button = product_card.find_element(
        By.XPATH,
        ".//button[contains(@onclick, 'cart.add')]"
    )

    driver.execute_script(
        "arguments[0].click();",
        add_to_cart_button
    )

    time.sleep(2)

    # Handle possible alert
    handle_alert(driver)

    print("Product added to cart.")

    take_screenshot(
        driver,
        "04_product_added.png"
    )

    # -----------------------------------------------------
    # 5. OPEN CART
    # -----------------------------------------------------

    print("\n5. Opening shopping cart...")

    cart_link = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//a[contains(@href, 'checkout/cart')]"
            )
        )
    )

    cart_link.click()

    wait.until(
        EC.url_contains("checkout/cart")
    )

    print("Cart opened.")

    take_screenshot(
        driver,
        "05_cart_before_quantity_update.png"
    )

    # -----------------------------------------------------
    # 6. VERIFY PRODUCT EXISTS IN CART
    # -----------------------------------------------------

    print("\n6. Verifying product in cart...")

    cart_product = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                f"//a[contains(@href, 'product_id=') and contains(normalize-space(), '{product}')]"
            )
        )
    )

    # Get the product name using text or value/attribute
    product_name = cart_product.get_attribute("textContent").strip()

    print("Product found in cart:", product_name)

    assert product.lower() in product_name.lower()

    print("Product verification successful.")

    print("Product verification successful.")

    # -----------------------------------------------------
    # 7. UPDATE QUANTITY
    # -----------------------------------------------------

    print("\n7. Updating quantity to:", quantity)

    quantity_field = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//input[contains(@name,'quantity')]"
            )
        )
    )

    quantity_field.click()

    quantity_field.send_keys(
        Keys.CONTROL + "a"
    )

    quantity_field.send_keys(
        str(quantity)
    )

    # Click Update button
    update_button = driver.find_element(
        By.XPATH,
        "//button[contains(@data-original-title,'Update')]"
    )

    update_button.click()

    time.sleep(2)

    handle_alert(driver)

    print("Quantity updated.")

    take_screenshot(
        driver,
        "06_quantity_updated.png"
    )

    # -----------------------------------------------------
    # 8. VERIFY CART DETAILS
    # -----------------------------------------------------

    print("\n8. Verifying cart details...")

    updated_quantity = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//input[contains(@name,'quantity')]"
            )
        )
    )

    actual_quantity = int(
        updated_quantity.get_attribute("value")
    )

    assert actual_quantity == quantity

    print(
        f"Quantity verification successful: {actual_quantity}"
    )

    # Verify cart contains product
    assert product.lower() in driver.page_source.lower()

    print("Cart product verification successful.")

    take_screenshot(
        driver,
        "07_cart_verification.png"
    )

    # -----------------------------------------------------
    # 9. HANDLE POPUPS / ALERTS
    # -----------------------------------------------------

    print("\n9. Checking for popup/alert...")

    handle_alert(driver)

    print("Popup/alert handling completed.")

    # -----------------------------------------------------
    # 10. FINAL SCREENSHOT
    # -----------------------------------------------------

    take_screenshot(
        driver,
        "08_final_cart.png"
    )

    print("\n===================================")
    print("TEST COMPLETED SUCCESSFULLY")
    print("===================================")