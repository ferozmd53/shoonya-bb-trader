# get_auth.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyotp
import time
import xlwings as xw
import requests

def get_auth_code():
    """Function to get authentication code"""
    print("="*60)
    print("SHOONYA AUTHENTICATION")
    print("="*60)
    
    wb = xw.Book("symbols.xlsx")
    ws = wb.sheets["LOGIN"]
    
    CLIENT_ID = ws.range("B2").value
    USER_ID = ws.range("B3").value
    PASSWORD = ws.range("B4").value
    TOTP_SECRET = ws.range("B5").value
    
    # Get and save IP address
    try:
        response = requests.get('https://api.ipify.org', timeout=5)
        ip = response.text
        print(f"Your IP: {ip}")
        ws.range("C11").value = ip
        wb.save()
    except:
        pass
    
    print(f"Client ID: {CLIENT_ID}")
    print(f"User ID: {USER_ID}")
    
    LOGIN_URL = f"https://trade.shoonya.com/OAuthlogin/investor-entry-level/login?api_key={CLIENT_ID}&route_to={USER_ID}"
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    driver.get(LOGIN_URL)
    time.sleep(3)
    
    inputs = [x for x in driver.find_elements(By.TAG_NAME, "input") if x.is_displayed()]
    
    if len(inputs) >= 3:
        inputs[0].send_keys(USER_ID)
        inputs[1].send_keys(PASSWORD)
        otp = pyotp.TOTP(TOTP_SECRET).now()
        inputs[2].send_keys(otp)
        time.sleep(1)
    
    for b in driver.find_elements(By.TAG_NAME, "button"):
        if "LOGIN" in b.text.upper():
            driver.execute_script("arguments[0].click();", b)
            print("LOGIN CLICKED")
            break
    
    while True:
        url = driver.current_url
        if "#/?code=" in url:
            code = url.split("code=")[1]
            ws.range("B7").value = code
            wb.save()
            print(f"\n✅ AUTH CODE: {code}")
            driver.quit()
            return code
        time.sleep(0.2)

if __name__ == "__main__":
    get_auth_code()
