from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

driver = None
FIREFOX_PROFILE_PATH = r"C:\Users\yaswa\AppData\Roaming\Mozilla\Firefox\Profiles\nc38qhq6.default"

def open_gmail():
    global driver
    try:
        profile = webdriver.FirefoxProfile(FIREFOX_PROFILE_PATH)
        options = Options()
        options.profile = profile  # Use the FirefoxProfile object
        driver = webdriver.Firefox(options=options)
    except Exception as e:
        return f"Error initializing Firefox WebDriver for Gmail: {e}"
    
    driver.get("https://mail.google.com")
    
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.zA"))
        )
    except Exception as e:
        return f"Error: Gmail did not load in time: {e}"
    
    return "Gmail opened with your real Firefox profile."
# first mail
base_x = 855
base_y = 335
# compose button
compose_x = 42
compose_y = 225
# send button
send_x = 1145
send_y = 1095
# Test the function if run directly:
if __name__ == "__main__":
    print(open_gmail())
