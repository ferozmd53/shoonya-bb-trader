"""
Main Bollinger Bands Trading System
"""

from NorenRestApiPy.NorenApi import NorenApi
import time
import datetime
from datetime import datetime, timedelta
import numpy as np
import json
import threading
from collections import deque
import xlwings as xw
import warnings
warnings.filterwarnings('ignore')

class BollingerBandsTrader:
    def __init__(self, excel_file="symbols.xlsx", sheet_name="symbols"):
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.excel_name = None
        self.api = None
        self.feed_opened = False
        self.live_data = {}
        self.symbol_tokens = {}
        self.token_symbols = {}
        self.indicators = {}
        self.historical_data_cache = {}
        self.tick_count = 0
        self.bb_period = 10
        self.bb_std = 1.8
        
    def login(self):
        try:
            self.excel_name = xw.Book(self.excel_file)
            class ShoonyaApiPy(NorenApi):
                def __init__(self):
                    super().__init__(host='https://api.shoonya.com/NorenWClientAPI/', 
                                   websocket='wss://api.shoonya.com/NorenWSAPI/')
            self.api = ShoonyaApiPy()
            login_sheet = self.excel_name.sheets['LOGIN']
            userid = login_sheet.range('B3').value
            api_secret = login_sheet.range('B6').value
            auth_code = login_sheet.range('B7').value
            if userid:
                userid = str(userid).strip()
            if api_secret:
                api_secret = str(api_secret).strip()
            if auth_code:
                auth_code = str(auth_code).strip()
            cred = {'client_id': f'{userid}_U', 'secret': api_secret, 'uid': userid}
            result = self.api.getAccessToken(auth_code, api_secret, cred['client_id'], userid)
            if result:
                acc_tok, usrid, ref_tok, actid = result
                self.api.injectOAuthHeader(acc_tok, userid, actid)
                print("Login Successful!")
                return True
        except Exception as e:
            print(f"Login error: {e}")
        return False
    
    def run(self):
        print("\n" + "="*80)
        print("Bollinger Bands Trading System")
        print("="*80)
        if not self.login():
            print("Login failed!")
            return
        print("System ready!")

def main():
    trader = BollingerBandsTrader()
    trader.run()

if __name__ == "__main__":
    main()
