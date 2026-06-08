"""
Shoonya API OAuth Authentication
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyotp
import time
import xlwings as xw
import json

class ShoonyaAuth:
    def __init__(self):
        self.client_id = None
        self.user_id = None
        self.password = None
        self.totp_secret = None
        self.secret_code = None
        self.auth_code = None
        
    def set_credentials(self, client_id, user_id, password, totp_secret, secret_code):
        self.client_id = client_id
        self.user_id = user_id
        self.password = password
        self.totp_secret = totp_secret
        self.secret_code = secret_code
        return self
    
    @classmethod
    def from_excel(cls, excel_file="symbols.xlsx", sheet_name="LOGIN"):
        try:
            wb = xw.Book(excel_file)
            ws = wb.sheets[sheet_name]
            auth = cls()
            auth.client_id = str(ws.range('B2').value).strip() if ws.range('B2').value else None
            auth.user_id = str(ws.range('B3').value).strip() if ws.range('B3').value else None
            auth.password = str(ws.range('B4').value).strip() if ws.range('B4').value else None
            auth.totp_secret = str(ws.range('B5').value).strip() if ws.range('B5').value else None
            auth.secret_code = str(ws.range('B6').value).strip() if ws.range('B6').value else None
            return auth
        except Exception as e:
            print(f"Error: {e}")
            return None
    
    def get_auth_code(self, headless=False):
        if not all([self.client_id, self.user_id, self.password, self.totp_secret]):
            print("Missing credentials!")
            return None
        
        login_url = f"https://trade.shoonya.com/OAuthlogin/investor-entry-level/login?api_key={self.client_id}&route_to={self.user_id}"
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument('--headless')
        
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.maximize_window()
        driver.get(login_url)
        time.sleep(3)
        
        inputs = [x for x in driver.find_elements(By.TAG_NAME, "input") if x.is_displayed()]
        if len(inputs) >= 3:
            inputs[0].send_keys(self.user_id)
            inputs[1].send_keys(self.password)
            otp = pyotp.TOTP(self.totp_secret).now()
            inputs[2].send_keys(otp)
            time.sleep(1)
        
        for button in driver.find_elements(By.TAG_NAME, "button"):
            if "LOGIN" in button.text.upper():
                driver.execute_script("arguments[0].click();", button)
                print("Login clicked")
                break
        
        auth_code = None
        for _ in range(30):
            url = driver.current_url
            if "#/?code=" in url:
                auth_code = url.split("code=")[1]
                print(f"Auth Code: {auth_code}")
                break
            time.sleep(0.2)
        
        driver.quit()
        self.auth_code = auth_code
        return auth_code
    
    def save_auth_to_excel(self, excel_file="symbols.xlsx", sheet_name="LOGIN"):
        if self.auth_code:
            try:
                wb = xw.Book(excel_file)
                ws = wb.sheets[sheet_name]
                ws.range('B7').value = self.auth_code
                wb.save()
                print(f"Saved to {excel_file}")
            except Exception as e:
                print(f"Error: {e}")

def main():
    print("\n" + "="*50)
    print("Shoonya OAuth Authentication")
    print("="*50)
    print("\nEnter your credentials:")
    client_id = input("Client ID (e.g., FA78603_U): ").strip()
    user_id = input("User ID (e.g., FA78603): ").strip()
    password = input("Password: ").strip()
    totp_secret = input("TOTP Secret (32 characters): ").strip()
    secret_code = input("Secret Code: ").strip()
    auth = ShoonyaAuth()
    auth.set_credentials(client_id, user_id, password, totp_secret, secret_code)
    print("\nGetting auth code...")
    code = auth.get_auth_code()
    if code:
        print(f"\nSuccess! Auth Code: {code}")
        save = input("\nSave to Excel? (y/n): ").lower()
        if save == 'y':
            auth.save_auth_to_excel()
    else:
        print("\nFailed to get auth code")

if __name__ == "__main__":
    main()
