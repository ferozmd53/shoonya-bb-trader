# get_auth.py - OAuth helper

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyotp
import time
import xlwings as xw

def get_auth_code():
    """Get authentication code from Shoonya"""
    print("="*60)
    print("SHOONYA AUTHENTICATION")
    print("="*60)
    
    # Open Excel
    wb = xw.Book("symbols.xlsx")
    ws = wb.sheets["LOGIN"]
    
    # Read credentials from Excel
    CLIENT_ID = ws.range("B2").value
    USER_ID = ws.range("B3").value
    PASSWORD = ws.range("B4").value
    TOTP_SECRET = ws.range("B5").value
    SECRET_CODE = ws.range("B6").value
    
    print(f"Client ID: {CLIENT_ID}")
    print(f"User ID: {USER_ID}")
    print("Getting auth code...")
    
    # Login URL
    LOGIN_URL = f"https://trade.shoonya.com/OAuthlogin/investor-entry-level/login?api_key={CLIENT_ID}&route_to={USER_ID}"
    
    # Open browser
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    driver.get(LOGIN_URL)
    time.sleep(3)
    
    # Fill credentials
    inputs = [x for x in driver.find_elements(By.TAG_NAME, "input") if x.is_displayed()]
    if len(inputs) >= 3:
        inputs[0].send_keys(USER_ID)
        inputs[1].send_keys(PASSWORD)
        otp = pyotp.TOTP(TOTP_SECRET).now()
        inputs[2].send_keys(otp)
        time.sleep(1)
    
    # Click login
    for b in driver.find_elements(By.TAG_NAME, "button"):
        if "LOGIN" in b.text.upper():
            driver.execute_script("arguments[0].click();", b)
            print("Login clicked")
            break
    
    # Get auth code
    auth_code = None
    for _ in range(30):
        url = driver.current_url
        if "#/?code=" in url:
            auth_code = url.split("code=")[1]
            ws.range("B7").value = auth_code
            wb.save()
            print(f"\n✅ AUTH CODE: {auth_code}")
            print("✅ Saved to Excel (B7)")
            break
        time.sleep(0.2)
    
    driver.quit()
    return auth_code

def main():
    get_auth_code()

if __name__ == "__main__":
    main()
