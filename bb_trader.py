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

# ============================================
# GLOBAL VARIABLES
# ============================================

excel_name = xw.Book('symbols.xlsx')
api = None
feed_opened = False
live_data = {}
symbol_tokens = {}
token_symbols = {}
indicators = {}
historical_data_cache = {}
tick_count = 0
last_symbol_check = 0
last_excel_update = 0

# ============================================
# CONFIGURATION
# ============================================

class Config:
    BB_PERIOD = 5
    BB_STD = 1.8
    EXCEL_UPDATE_INTERVAL = 0.2

# ============================================
# SAFE CONVERSION
# ============================================

def safe_float(value, default=0):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value)
        except:
            return default
    return default

def safe_int(value, default=0):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        try:
            return int(float(value))
        except:
            return default
    return default

# ============================================
# BOLLINGER BANDS
# ============================================

class IncrementalBollingerBands:
    def __init__(self, period=10, num_std=1.8, symbol=""):
        self.period = period
        self.num_std = num_std
        self.symbol = symbol
        self.daily_closes = []
        self.today_ltp = 0
        self.sma = None
        self.upper = None
        self.lower = None
        self.initialized = False
        
    def set_daily_closes(self, closes):
        self.daily_closes = closes.copy()
        
    def add_today_price(self, price):
        if price <= 0:
            return
        self.today_ltp = price
        self.calculate_bb()
        
    def calculate_bb(self):
        if len(self.daily_closes) >= self.period - 1 and self.today_ltp > 0:
            hist_closes = self.daily_closes[-(self.period - 1):]
            all_prices = hist_closes + [self.today_ltp]
            
            self.sma = sum(all_prices) / self.period
            variance = sum((x - self.sma) ** 2 for x in all_prices) / self.period
            std = variance ** 0.5
            self.upper = self.sma + (self.num_std * std)
            self.lower = self.sma - (self.num_std * std)
            self.initialized = True
            
    def get_bands(self):
        return self.sma, self.upper, self.lower

# ============================================
# API CLASS
# ============================================

class ShoonyaApiPy(NorenApi):
    def __init__(self):
        super().__init__(
            host='https://api.shoonya.com/NorenWClientAPI/',
            websocket='wss://api.shoonya.com/NorenWSAPI/'
        )

# ============================================
# LOGIN
# ============================================

def Shoonya_login():
    global api
    try:     
        class ShoonyaApiPy(NorenApi):
            def __init__(self):
                super().__init__(host='https://api.shoonya.com/NorenWClientAPI/', 
                               websocket='wss://api.shoonya.com/NorenWSAPI/')

        api = ShoonyaApiPy()
        
        try:
            login_sheet = excel_name.sheets['LOGIN']
            userid = login_sheet.range('B3').value
            api_secret = login_sheet.range('B6').value
            auth_code = login_sheet.range('B7').value
            
            if not userid or not api_secret or not auth_code:
                print("❌ Missing credentials in LOGIN sheet!")
                print("   Please ensure cells B3, B6, and B7 have values")
                return 0
                
            userid = str(userid).strip()
            api_secret = str(api_secret).strip()
            auth_code = str(auth_code).strip()
            
            print(f"   Using credentials from LOGIN sheet")
            
        except Exception as e:
            print(f"❌ Error reading LOGIN sheet: {e}")
            print("   Please make sure you have a sheet named 'LOGIN'")
            print("   with credentials in B3 (User ID), B6 (API Secret), and B7 (Auth Code)")
            return 0
        
        cred = {'client_id': f'{userid}_U', 'secret': api_secret, 'uid': userid}
        result = api.getAccessToken(auth_code, api_secret, cred['client_id'], userid)
#        print(result)
        if result:
            acc_tok, usrid, ref_tok, actid = result
#            login_sheet = excel_name.sheets['LOGIN']
            acess_token = login_sheet.range('B9').value =     acc_tok       
            user_token = login_sheet.range('B10').value =    ref_tok
            print("✅ TOKEN")
            api.injectOAuthHeader(acc_tok, userid, actid)
            print("✅ Login Successful!")
            return 1
        else:
            print("❌ Login failed - Invalid credentials or network issue")
            
    except Exception as e:
        print(f"Login error: {e}")
    return 0

# ============================================
# GET TOKEN
# ============================================

def GetToken(exchange, tradingsymbol):
    try:
        search = tradingsymbol.replace('-EQ', '').strip()
        result = api.searchscrip(exchange=exchange, searchtext=search)
        if result and result.get('values'):
            for item in result['values']:
                tsym = item.get('tsym', '').upper()
                if tsym in [tradingsymbol.upper(), search.upper()]:
                    return item.get('token')
            return result['values'][0].get('token')
    except Exception as e:
        print(f"Token error: {e}")
    return None

# ============================================
# FETCH HISTORICAL DATA
# ============================================

def fetch_historical_data(symbol):
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=40)
        
        start_epoch = int(start_date.timestamp())
        end_epoch = int(end_date.timestamp())
        
        ret = api.get_daily_price_series(
            exchange="NSE", tradingsymbol=symbol,
            startdate=str(start_epoch), enddate=str(end_epoch)
        )
        
        if not ret:
            return None
        
        parsed_data = []
        for item in ret:
            if isinstance(item, str):
                try:
                    parsed_data.append(json.loads(item))
                except:
                    continue
        
        if not parsed_data:
            return None
        
        import pandas as pd
        df = pd.DataFrame(parsed_data)
        df.rename(columns={
            'time': 'datetime', 'into': 'open', 'inth': 'high',
            'intl': 'low', 'intc': 'close', 'intv': 'volume'
        }, inplace=True)
        
        df['datetime'] = pd.to_datetime(df['datetime'], format='%d-%b-%Y')
        for col in ['open', 'high', 'low', 'close', 'volume']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        df = df.sort_values('datetime')
        
        if len(df) < Config.BB_PERIOD + 1:
            return None
        
        daily_closes = df['close'].tolist()
        yesterday = df.iloc[-1]
        
        last_10_closes = df.iloc[-Config.BB_PERIOD:]['close'].tolist()
        sma_y = sum(last_10_closes) / Config.BB_PERIOD
        variance_y = sum((x - sma_y) ** 2 for x in last_10_closes) / Config.BB_PERIOD
        std_y = variance_y ** 0.5
        upper_y = sma_y + (Config.BB_STD * std_y)
        lower_y = sma_y - (Config.BB_STD * std_y)
        
        yesterday_data = {
            'date': yesterday['datetime'].strftime('%d/%m/%Y'),
            'open': yesterday['open'],
            'high': yesterday['high'],
            'low': yesterday['low'],
            'close': yesterday['close'],
            'volume': yesterday['volume'],
            'bb_upper': upper_y,
            'bb_middle': sma_y,
            'bb_lower': lower_y
        }
        
        return {
            'daily_closes': daily_closes,
            'yesterday': yesterday_data
        }
        
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
        return None

# ============================================
# WEBSOCKET CALLBACKS
# ============================================

def on_ticks(tick):
    global live_data, indicators, tick_count
    
    try:
        if isinstance(tick, str):
            tick = json.loads(tick)
        
        tick_count += 1
        
        key = f"{tick['e']}|{tick['tk']}"
        
        if key in live_data:
            d = live_data[key]
            ltp = safe_float(tick.get('lp', d.get('ltp', 0)))
            
            d['ltp'] = ltp
            d['open'] = safe_float(tick.get('o', d.get('open', 0)))
            d['high'] = safe_float(tick.get('h', d.get('high', 0)))
            d['low'] = safe_float(tick.get('l', d.get('low', 0)))
            d['close'] = safe_float(tick.get('c', d.get('close', 0)))
            d['volume'] = safe_int(tick.get('v', d.get('volume', 0)))
            d['timestamp'] = datetime.now()
            
            if d.get('first_tick', True):
                d['first_tick'] = False
                print(f"\n✓ First tick for {d['symbol']}: LTP={ltp}")
            
            symbol = d.get('symbol', '')
            if symbol and symbol in indicators:
                indicators[symbol].add_today_price(ltp)
                sma, upper, lower = indicators[symbol].get_bands()
                d['bb_upper'] = upper
                d['bb_middle'] = sma
                d['bb_lower'] = lower
                
                if tick_count <= 10:
                    print(f"   Today's BB for {symbol}: Upper={upper:.2f}, Middle={sma:.2f}, Lower={lower:.2f}")
    except:
        pass

def on_open():
    global feed_opened
    feed_opened = True
    print("✅ WebSocket Connected")

def on_close():
    global feed_opened
    feed_opened = False
    print("❌ WebSocket Closed")

def on_order(order):
    pass

# ============================================
# SUBSCRIBE SYMBOLS
# ============================================

def subscribe_symbols(tokens_list):
    if not tokens_list:
        return
    for i in range(0, len(tokens_list), 10):
        try:
            api.subscribe(tokens_list[i:i+10])
            print(f"✓ Subscribed to {len(tokens_list[i:i+10])} symbols")
        except Exception as e:
            print(f"Subscribe error: {e}")
        time.sleep(0.1)

# ============================================
# CHECK NEW SYMBOLS
# ============================================

def check_new_symbols():
    global last_symbol_check, indicators
    
    try:
        ws = excel_name.sheets['symbols']
        symbols_data = ws.range("A2:A100").value
        current = [str(s).strip().upper() for s in symbols_data if s] if symbols_data else []
        
        cleaned = []
        for s in current:
            s_str = s.upper()
            if s_str.startswith('NSE:'):
                s_str = s_str[4:]
            if not s_str.endswith('-EQ'):
                s_str = f"{s_str}-EQ"
            cleaned.append(s_str)
        
        new_symbols = [s for s in cleaned if s not in symbol_tokens]
        
        if new_symbols:
            print(f"\n🆕 Found {len(new_symbols)} new symbols")
            new_tokens = []
            for symbol in new_symbols:
                try:
                    token = GetToken("NSE", symbol)
                    if token:
                        tk = f"NSE|{token}"
                        symbol_tokens[symbol] = tk
                        token_symbols[token] = symbol
                        live_data[tk] = {
                            'ltp': 0, 'open': 0, 'high': 0, 'low': 0, 'close': 0,
                            'volume': 0, 'symbol': symbol, 'first_tick': True,
                            'timestamp': None, 'bb_upper': 0, 'bb_middle': 0, 'bb_lower': 0
                        }
                        indicators[symbol] = IncrementalBollingerBands(Config.BB_PERIOD, Config.BB_STD, symbol)
                        new_tokens.append(tk)
                        
                        hist_data = fetch_historical_data(symbol)
                        if hist_data:
                            indicators[symbol].set_daily_closes(hist_data['daily_closes'])
                            historical_data_cache[symbol] = hist_data['yesterday']
                        
                        print(f"   ✓ Added {symbol}")
                except Exception as e:
                    print(f"   Error adding {symbol}: {e}")
                time.sleep(0.05)
            
            if new_tokens and feed_opened:
                subscribe_symbols(new_tokens)
                print(f"✓ Subscribed to {len(new_tokens)} new symbols\n")
    except Exception as e:
        pass

# ============================================
# UPDATE EXCEL - HISTORICAL DATA AT COLUMN M
# ============================================

def update_excel_bulk():
    global last_excel_update
    
    try:
        current_time = time.time()
        if current_time - last_excel_update < Config.EXCEL_UPDATE_INTERVAL:
            return
        last_excel_update = current_time
        
        ws = excel_name.sheets['symbols']
        symbols_data = ws.range("A2:A200").value
        if not symbols_data:
            return
        
        live_rows = []
        hist_rows = []
        
        for symbol_cell in symbols_data:
            if not symbol_cell:
                live_rows.append([''] * 10)
                hist_rows.append([''] * 9)
                continue
            
            symbol = str(symbol_cell).strip().upper()
            if symbol.startswith('NSE:'):
                symbol = symbol[4:]
            if not symbol.endswith('-EQ'):
                symbol = f"{symbol}-EQ"
            
            tk = symbol_tokens.get(symbol)
            
            if tk and tk in live_data:
                d = live_data[tk]
                ltp = d.get('ltp', 0)
                bb_upper = d.get('bb_upper', 0)
                bb_middle = d.get('bb_middle', 0)
                bb_lower = d.get('bb_lower', 0)
                
                # Live data row - 10 columns (B to K)
                live_rows.append([
                    ltp if ltp > 0 else '',
                    d.get('open', 0) if d.get('open', 0) > 0 else '',
                    d.get('high', 0) if d.get('high', 0) > 0 else '',
                    d.get('low', 0) if d.get('low', 0) > 0 else '',
                    d.get('close', 0) if d.get('close', 0) > 0 else '',
                    d.get('volume', 0) if d.get('volume', 0) > 0 else '',
                    round(bb_upper, 2) if bb_upper else '',
                    round(bb_middle, 2) if bb_middle else '',
                    round(bb_lower, 2) if bb_lower else '',
                    d['timestamp'].strftime('%H:%M:%S') if d['timestamp'] else ''
                ])
                
                # Historical data - 9 columns (M to U)
                hist = historical_data_cache.get(symbol)
                if hist:
                    hist_rows.append([
                        hist['date'],
                        hist['open'] if hist['open'] > 0 else '',
                        hist['high'] if hist['high'] > 0 else '',
                        hist['low'] if hist['low'] > 0 else '',
                        hist['close'] if hist['close'] > 0 else '',
                        hist['volume'] if hist['volume'] > 0 else '',
                        round(hist['bb_upper'], 2) if hist['bb_upper'] else '',
                        round(hist['bb_middle'], 2) if hist['bb_middle'] else '',
                        round(hist['bb_lower'], 2) if hist['bb_lower'] else ''
                    ])
                else:
                    hist_rows.append([''] * 9)
            else:
                live_rows.append([''] * 10)
                hist_rows.append([''] * 9)
        
        # Update Excel with error handling
        try:
            if live_rows:
                # Live data: B2 to K (columns 2 to 11)
                ws.range(f"B2:K{2 + len(live_rows) - 1}").value = live_rows
            
            if hist_rows:
                # Historical data: M2 to U (columns 13 to 21)
                ws.range(f"M2:U{2 + len(hist_rows) - 1}").value = hist_rows
                
        except Exception as write_error:
            # Silently skip if Excel is in edit mode
            if "Member not found" not in str(write_error) and "-2147352573" not in str(write_error):
                print(f"Excel write error: {write_error}")
            
    except Exception as e:
        if "Member not found" not in str(e) and "-2147352573" not in str(e):
            print(f"Excel update error: {e}")

# ============================================
# SETUP HEADERS - HISTORICAL AT COLUMN M
# ============================================

def setup_excel_headers():
    try:
        ws = excel_name.sheets['symbols']
        
        # Clear existing headers
        ws.range("1:1").clear_contents()
        
        # SYMBOL header (Column A)
        ws.range('A1').value = 'Symbol'
        ws.range('A1').color = (54, 96, 146)
        ws.range('A1').font.color = (255, 255, 255)
        ws.range('A1').font.bold = True
        
        # LIVE DATA HEADERS - Columns B to K (10 columns)
        # B=LTP, C=Open, D=High, E=Low, F=Close, G=Volume, H=BB Upper, I=BB Middle, J=BB Lower, K=Last Update
        live_headers = ['LTP', 'Open', 'High', 'Low', 'Close', 'Volume',
                        'BB Upper', 'BB Middle', 'BB Lower', 'Last Update']
        
        for col_idx, header in enumerate(live_headers, start=2):
            cell = ws.range((1, col_idx))
            cell.value = header
            cell.color = (54, 96, 146)
            cell.font.color = (255, 255, 255)
            cell.font.bold = True
        
        # HISTORICAL DATA HEADERS - Columns M to U (9 columns)
        # M=Date, N=Open, O=High, P=Low, Q=Close, R=Volume, S=BB Upper, T=BB Middle, U=BB Lower
        hist_headers = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume',
                        'BB Upper', 'BB Middle', 'BB Lower']
        
        for col_idx, header in enumerate(hist_headers, start=13):
            cell = ws.range((1, col_idx))
            cell.value = header
            cell.color = (146, 96, 54)
            cell.font.color = (255, 255, 255)
            cell.font.bold = True
        
        # Clear any old headers in between
        ws.range('L1').value = ''  # Column L is empty separator
        
        # Set column widths
        ws.range('A:A').column_width = 20      # Symbol
        ws.range('B:K').column_width = 12      # Live data (10 columns)
        ws.range('L:L').column_width = 3       # Separator
        ws.range('M:U').column_width = 12      # Historical data (9 columns)

        return True
    except Exception as e:
        print(f"Error setting up headers: {e}")
        return False

# ============================================
# READ SYMBOLS
# ============================================

def read_symbols_from_excel():
    try:
        ws = excel_name.sheets['symbols']
        symbols_data = ws.range("A2:A200").value
        symbols = [str(s).strip().upper() for s in symbols_data if s] if symbols_data else []
        
        cleaned = []
        seen = set()
        for s in symbols:
            s_str = s.upper()
            if s_str.startswith('NSE:'):
                s_str = s_str[4:]
            if not s_str.endswith('-EQ'):
                s_str = f"{s_str}-EQ"
            if s_str not in seen:
                seen.add(s_str)
                cleaned.append(s_str)
        return cleaned
    except Exception as e:
        return []

# ============================================
# MAIN EXCEL LOOP
# ============================================

def start_excel_loop():
    global last_symbol_check
    
    print("✓ Starting REAL-TIME Excel update loop...\n")
    
    update_count = 0
    last_status = time.time()
    
    while True:
        try:
            current = time.time()
            
            if current - last_symbol_check >= 5:
                check_new_symbols()
                last_symbol_check = current
            
            update_excel_bulk()
            
            update_count += 1
            if update_count % 30 == 0:
                try:
                    excel_name.save()
                except:
                    pass
            
            if current - last_status >= 10:
                active = sum(1 for d in live_data.values() if d.get('ltp', 0) > 0)
                print(f"📈 Active: {active}/{len(symbol_tokens)} | Ticks: {tick_count}")
                tick_count = 0
                last_status = current
            
            time.sleep(0.05)
            
        except Exception as e:
            time.sleep(0.1)

# ============================================
# MAIN
# ============================================

def main():
    global indicators, historical_data_cache
    
    print("\n" + "="*80)
    print("🚀 BOLLINGER BANDS TRADING SYSTEM")
    print("="*80)
    
    print("\n[1/4] Setting up Excel...")
    setup_excel_headers()
    
    print("\n[2/4] Logging to Shoonya...")
    if not Shoonya_login():
        print("❌ Login failed! Check your LOGIN sheet in Excel.")
        return
    
    print("\n[3/4] Reading symbols from Column A...")
    symbols = read_symbols_from_excel()
    
    if not symbols:
        default = ["RELIANCE-EQ", "TCS-EQ", "INFY-EQ"]
        print(f"⚠️ No symbols found, adding default symbols to column A")
        for i, sym in enumerate(default, start=2):
            excel_name.sheets['symbols'].range(f"A{i}").value = sym
        excel_name.save()
        symbols = default
        print(f"✅ Added default symbols")
    
    print(f"📋 Processing {len(symbols)} symbols...")
    
    print("\n[4/4] Initializing symbols...")
    tokens = []
    for i, symbol in enumerate(symbols, 1):
        print(f"   [{i:2}/{len(symbols)}] {symbol}...", end=" ")
        token = GetToken("NSE", symbol)
        if token:
            tk = f"NSE|{token}"
            symbol_tokens[symbol] = tk
            token_symbols[token] = symbol
            live_data[tk] = {
                'ltp': 0, 'open': 0, 'high': 0, 'low': 0,
                'close': 0, 'volume': 0, 'symbol': symbol,
                'first_tick': True, 'timestamp': None,
                'bb_upper': 0, 'bb_middle': 0, 'bb_lower': 0
            }
            indicators[symbol] = IncrementalBollingerBands(Config.BB_PERIOD, Config.BB_STD, symbol)
            tokens.append(tk)
            
            hist_data = fetch_historical_data(symbol)
            if hist_data:
                indicators[symbol].set_daily_closes(hist_data['daily_closes'])
                historical_data_cache[symbol] = hist_data['yesterday']
                print(f"✓ (Historical: {hist_data['yesterday']['date']})")
            else:
                print(f"✓ (No historical data)")
        else:
            print(f"✗ FAILED - Token not found")
        time.sleep(0.03)
    
    print(f"\n✅ Initialized {len(tokens)} symbols")
    
    if len(tokens) == 0:
        print("❌ No symbols initialized! Exiting.")
        return
    
    print("\nStarting WebSocket connection...")
    
    try:
        api.start_websocket(
            subscribe_callback=on_ticks,
            order_update_callback=on_order,
            socket_open_callback=on_open,
            socket_close_callback=on_close
        )
        
        print("   Waiting for WebSocket to connect...")
        for _ in range(15):
            if feed_opened:
                break
            time.sleep(1)
            print(f"   ... waiting ({_+1}/15)")
        
        if feed_opened:
            print("✅ WebSocket connected!")
            if tokens:
                print(f"\n📡 Subscribing to {len(tokens)} symbols...")
                subscribe_symbols(tokens)
                print(f"✓ Subscribed to all symbols")
            print("\n🚀 Starting real-time data feed...")
            print("   Data updates every 0.2 seconds")
            print("   Column Layout:")
            print("     - Live Data: Columns B to K (LTP, Open, High, Low, Close, Volume, BB Upper, BB Middle, BB Lower, Last Update)")
            print("     - Historical: Columns M to U (Date, Open, High, Low, Close, Volume, BB Upper, BB Middle, BB Lower)")
            print("\n   NOTE: If you click on a cell, updates will pause until you press Enter\n")
            start_excel_loop()
        else:
            print("❌ WebSocket connection failed after 15 seconds!")
            
    except Exception as e:
        print(f"❌ Failed to start WebSocket: {e}")

if __name__ == "__main__":
    main()
