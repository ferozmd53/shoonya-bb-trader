# get_auth.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyotp
import time
import xlwings as xw

# Read from Excel - CORRECT CELLS
wb = xw.Book("symbols.xlsx")
ws = wb.sheets["LOGIN"]

CLIENT_ID = ws.range("B2").value      # FA78603_U
USER_ID = ws.range("B3").value        # FA78603
PASSWORD = ws.range("B4").value       # your password
TOTP_SECRET = ws.range("B5").value    # your 32-char secret
SECRET_CODE = ws.range("B6").value    # your secret code

# Check if credentials exist
if not CLIENT_ID or not USER_ID or not PASSWORD or not TOTP_SECRET:
    print("❌ Missing credentials in Excel!")
    print("Please fill B2, B3, B4, B5 in LOGIN sheet")
    exit()

print(f"Client ID: {CLIENT_ID}")
print(f"User ID: {USER_ID}")

# Get IP
try:
    import requests
    response = requests.get('https://api.ipify.org', timeout=5)
    print(f"Your IP: {response.text}")
except:
    pass

# Login URL
LOGIN_URL = f"https://trade.shoonya.com/OAuthlogin/investor-entry-level/login?api_key={CLIENT_ID}&route_to={USER_ID}"

# Open Chrome
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

# Click Login
for b in driver.find_elements(By.TAG_NAME, "button"):
    if "LOGIN" in b.text.upper():
        driver.execute_script("arguments[0].click();", b)
        print("LOGIN CLICKED")
        break

# Get Auth Code
while True:
    url = driver.current_url
    print(url)
    if "#/?code=" in url:
        code = url.split("code=")[1]
        ws.range("B7").value = code
        wb.save()
        print(f"\n✅ AUTH CODE: {code}")
        print("✅ Saved to Excel (B7)")
        break
    time.sleep(0.2)

driver.quit()
print("\n✅ Authentication complete!")
