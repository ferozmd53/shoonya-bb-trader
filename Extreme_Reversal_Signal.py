# Extreme_Reversal_Signal.py - StochRSI + Bollinger Bands (Full Working Code)
from NorenRestApiPy.NorenApi import NorenApi
import time
import datetime
from datetime import datetime, timedelta
import numpy as np
import json
import xlwings as xw
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# ====================================================================
# GLOBAL VARIABLES
# ====================================================================

excel_name = xw.Book('symbols.xlsx')
api = None
feed_opened = False
live_data = {}
symbol_tokens = {}
token_symbols = {}
historical_data_cache = {}
tick_count = 0
last_symbol_check = 0
last_excel_update = 0

# ====================================================================
# CONFIGURATION
# ====================================================================

class Config:
    MA_LENGTH = 20
    STD_UP = 2.0
    STD_DOWN = 2.0
    RSI_LENGTH = 14
    STO_LENGTH = 14
    STO_UPPER = 70
    STO_LOWER = 30
    EXCEL_UPDATE_INTERVAL = 1
    LOAD_DAYS = 500
    KEEP_DAYS = 100

# ====================================================================
# API CLASS
# ====================================================================

class ShoonyaApiPy(NorenApi):
    def __init__(self):
        super().__init__(
            host='https://api.shoonya.com/NorenWClientAPI/',
            websocket='wss://api.shoonya.com/NorenWSAPI/'
        )

# ====================================================================
# STOCHRSI INDICATOR CLASS
# ====================================================================

class StochRSIIndicator:
    def __init__(self, symbol, all_closes):
        self.symbol = symbol
        self.all_closes = list(all_closes) if all_closes else []
        self.closes = self.all_closes[-Config.KEEP_DAYS:] if len(self.all_closes) > Config.KEEP_DAYS else self.all_closes.copy()
        
        self.avg_gain = 0
        self.avg_loss = 0
        self.rsi_series = []
        self.stoch_series = []
        
        self.prev_close = None
        self.prev_bb_lower = 0
        self.prev_bb_upper = 0
        self.prev_stoch = 50
        
        self.current_rsi = 50
        self.current_stoch = 50
        self.current_sma_stoch = 50
        self.current_bb_middle = 0
        self.current_bb_upper = 0
        self.current_bb_lower = 0
        
        if len(self.all_closes) >= Config.MA_LENGTH + Config.RSI_LENGTH:
            self._initialize_indicators()
    
    def _initialize_indicators(self):
        df = pd.DataFrame({'close': self.all_closes})
        delta = df['close'].diff()
        u = delta.where(delta > 0, 0.0)
        d = -delta.where(delta < 0, 0.0)
        
        alpha = 1.0 / Config.RSI_LENGTH
        avg_gain = u.ewm(alpha=alpha, adjust=False).mean()
        avg_loss = d.ewm(alpha=alpha, adjust=False).mean()
        
        rs = avg_gain / avg_loss
        rsi = 100 - (100 / (1 + rs))
        
        rsi_values = rsi.dropna().tolist()
        self.rsi_series = rsi_values[-Config.KEEP_DAYS:] if len(rsi_values) > Config.KEEP_DAYS else rsi_values
        self.current_rsi = self.rsi_series[-1] if self.rsi_series else 50
        
        self.avg_gain = avg_gain.iloc[-1] if not pd.isna(avg_gain.iloc[-1]) else 0
        self.avg_loss = avg_loss.iloc[-1] if not pd.isna(avg_loss.iloc[-1]) else 0
        
        if len(rsi_values) >= Config.STO_LENGTH:
            rsi_array = np.array(rsi_values)
            stoch_values = []
            for i in range(Config.STO_LENGTH - 1, len(rsi_array)):
                window = rsi_array[i - Config.STO_LENGTH + 1:i + 1]
                lowest_rsi = np.min(window)
                highest_rsi = np.max(window)
                if highest_rsi != lowest_rsi:
                    stoch = 100 * (rsi_array[i] - lowest_rsi) / (highest_rsi - lowest_rsi)
                else:
                    stoch = 50
                stoch_values.append(stoch)
            
            self.stoch_series = stoch_values[-Config.KEEP_DAYS:] if len(stoch_values) > Config.KEEP_DAYS else stoch_values
            self.current_stoch = self.stoch_series[-1] if self.stoch_series else 50
            if len(self.stoch_series) >= 3:
                self.current_sma_stoch = sum(self.stoch_series[-3:]) / 3
                if len(self.stoch_series) >= 2:
                    self.prev_stoch = self.stoch_series[-2]
        
        recent_closes = self.closes[-Config.MA_LENGTH:]
        self.current_bb_middle = sum(recent_closes) / Config.MA_LENGTH
        variance = sum((x - self.current_bb_middle) ** 2 for x in recent_closes) / Config.MA_LENGTH
        std = variance ** 0.5
        self.current_bb_upper = self.current_bb_middle + (Config.STD_UP * std)
        self.current_bb_lower = self.current_bb_middle - (Config.STD_DOWN * std)
        
        if len(self.closes) >= Config.MA_LENGTH + 1:
            prev_closes = self.closes[-Config.MA_LENGTH-1:-1]
            prev_middle = sum(prev_closes) / Config.MA_LENGTH
            prev_variance = sum((x - prev_middle) ** 2 for x in prev_closes) / Config.MA_LENGTH
            prev_std = prev_variance ** 0.5
            self.prev_bb_lower = prev_middle - (Config.STD_DOWN * prev_std)
            self.prev_bb_upper = prev_middle + (Config.STD_UP * prev_std)
        
        print(f"   ✅ {self.symbol}: RSI={self.current_rsi:.1f}, Stoch={self.current_stoch:.1f}")
    
    def _update_rsi_wilder(self, new_price):
        if self.prev_close is None:
            return self.current_rsi
        delta = new_price - self.prev_close
        gain = delta if delta > 0 else 0
        loss = -delta if delta < 0 else 0
        alpha = 1.0 / Config.RSI_LENGTH
        self.avg_gain = self.avg_gain * (1 - alpha) + gain * alpha
        self.avg_loss = self.avg_loss * (1 - alpha) + loss * alpha
        if self.avg_loss == 0:
            return 100
        rs = self.avg_gain / self.avg_loss
        return 100 - (100 / (1 + rs))
    
    def _update_stoch(self, rsi):
        self.rsi_series.append(rsi)
        if len(self.rsi_series) > Config.KEEP_DAYS:
            self.rsi_series = self.rsi_series[-Config.KEEP_DAYS:]
        if len(self.rsi_series) >= Config.STO_LENGTH:
            window = self.rsi_series[-Config.STO_LENGTH:]
            lowest_rsi = min(window)
            highest_rsi = max(window)
            if highest_rsi != lowest_rsi:
                stoch = 100 * (rsi - lowest_rsi) / (highest_rsi - lowest_rsi)
            else:
                stoch = 50
            self.stoch_series.append(stoch)
            if len(self.stoch_series) > Config.KEEP_DAYS:
                self.stoch_series = self.stoch_series[-Config.KEEP_DAYS:]
            self.current_stoch = stoch
            if len(self.stoch_series) >= 3:
                self.prev_stoch = self.current_sma_stoch
                self.current_sma_stoch = sum(self.stoch_series[-3:]) / 3
            return stoch
        return 50
    
    def _update_bb(self, new_price):
        self.prev_bb_lower = self.current_bb_lower
        self.prev_bb_upper = self.current_bb_upper
        self.closes.append(new_price)
        if len(self.closes) > Config.KEEP_DAYS:
            self.closes = self.closes[-Config.KEEP_DAYS:]
        if len(self.closes) >= Config.MA_LENGTH:
            recent_closes = self.closes[-Config.MA_LENGTH:]
            self.current_bb_middle = sum(recent_closes) / Config.MA_LENGTH
            variance = sum((x - self.current_bb_middle) ** 2 for x in recent_closes) / Config.MA_LENGTH
            std = variance ** 0.5
            self.current_bb_upper = self.current_bb_middle + (Config.STD_UP * std)
            self.current_bb_lower = self.current_bb_middle - (Config.STD_DOWN * std)
    
    def add_tick(self, ltp):
        if len(self.closes) == 0:
            self.closes.append(ltp)
            self.prev_close = ltp
            return None
        
        self.prev_close = self.closes[-1]
        rsi = self._update_rsi_wilder(ltp)
        self.current_rsi = rsi
        self._update_stoch(rsi)
        self._update_bb(ltp)
        
        buy_signal = False
        sell_signal = False
        signal = ""
        
        if self.prev_bb_lower != 0:
            if (self.prev_close < self.prev_bb_lower and 
                ltp > self.current_bb_lower and 
                self.prev_stoch < Config.STO_LOWER):
                buy_signal = True
                signal = "BUY"
                print(f"🔵 BUY {self.symbol} | Price:{ltp:.2f}")
            elif (self.prev_close > self.prev_bb_upper and 
                  ltp < self.current_bb_upper and 
                  self.prev_stoch > Config.STO_UPPER):
                sell_signal = True
                signal = "SELL"
                print(f"🔴 SELL {self.symbol} | Price:{ltp:.2f}")
        
        return {
            'rsi': round(self.current_rsi, 2),
            'stoch_rsi': round(self.current_stoch, 2),
            'sma_stoch': round(self.current_sma_stoch, 2),
            'bb_upper': round(self.current_bb_upper, 2),
            'bb_middle': round(self.current_bb_middle, 2),
            'bb_lower': round(self.current_bb_lower, 2),
            'buy': 1 if buy_signal else '',
            'sell': 1 if sell_signal else '',
            'signal': signal
        }

# ====================================================================
# SAFE CONVERSION FUNCTIONS
# ====================================================================

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

# ====================================================================
# SHOONYA API FUNCTIONS
# ====================================================================

def Shoonya_login():
    global api
    try:     
        api = ShoonyaApiPy()
        try:
            login_sheet = excel_name.sheets['LOGIN']
            userid = login_sheet.range('B3').value
            secret = login_sheet.range('B6').value
            auth = login_sheet.range('B7').value
            if not userid or not secret or not auth:
                print("❌ Missing credentials in LOGIN sheet!")
                return 0
            userid = str(userid).strip()
            secret = str(secret).strip()
            auth = str(auth).strip()
        except Exception as e:
            print(f"❌ Error reading LOGIN sheet: {e}")
            return 0
        
        cred = {'client_id': f'{userid}_U', 'secret': secret, 'uid': userid}
        result = api.getAccessToken(auth, secret, cred['client_id'], userid)
        if result:
            acc_tok, usrid, ref_tok, actid = result
            login_sheet = excel_name.sheets['LOGIN']
            login_sheet.range('B9').value = acc_tok       
            login_sheet.range('B10').value = ref_tok
            print("✅ TOKEN")
            api.injectOAuthHeader(acc_tok, userid, actid)
            print("✅ Login Successful!")
            return 1
        else:
            print("❌ Login failed")
    except Exception as e:
        print(f"Login error: {e}")
    return 0

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

def fetch_historical_data(symbol):
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=Config.LOAD_DAYS)
        start_epoch = int(start_date.timestamp())
        end_epoch = int(end_date.timestamp())
        
        ret = api.get_daily_price_series(
            exchange="NSE", 
            tradingsymbol=symbol,
            startdate=str(start_epoch), 
            enddate=str(end_epoch)
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
        df = df.dropna().reset_index(drop=True)
        
        if len(df) < Config.MA_LENGTH:
            return None
        
        yesterday = df.iloc[-1]
        yesterday_data = {
            'date': yesterday['datetime'].strftime('%d/%m/%Y'),
            'open': yesterday['open'],
            'high': yesterday['high'],
            'low': yesterday['low'],
            'close': yesterday['close'],
            'volume': yesterday['volume'],
            'bb_upper': 0, 'bb_middle': 0, 'bb_lower': 0,
            'rsi': 0, 'stoch_rsi': 0, 'sma_stoch': 0,
            'buy': '', 'sell': '', 'signal': ''
        }
        
        return {
            'all_closes': df['close'].tolist(),
            'yesterday': yesterday_data
        }
        
    except Exception as e:
        print(f"Error fetching {symbol}: {e}")
        return None

# ====================================================================
# WEBSOCKET CALLBACKS
# ====================================================================

def on_ticks(tick):
    global live_data, tick_count
    try:
        if isinstance(tick, str):
            tick = json.loads(tick)
        
        tick_count += 1
        key = f"{tick['e']}|{tick['tk']}"
        
        if key in live_data:
            d = live_data[key]
            ltp = safe_float(tick.get('lp', d.get('ltp', 0)))
            volume = safe_int(tick.get('v', d.get('volume', 0)))
            open_price = safe_float(tick.get('o', d.get('open', 0)))
            
            if d.get('first_tick', True):
                d['first_tick'] = False
                d['open'] = open_price if open_price > 0 else ltp
                d['high'] = ltp
                d['low'] = ltp
                print(f"\n✓ First tick for {d['symbol']}: LTP={ltp}")
            
            if ltp > d.get('high', 0):
                d['high'] = ltp
            if ltp < d.get('low', 999999):
                d['low'] = ltp
            
            d['ltp'] = ltp
            d['close'] = ltp
            d['volume'] = volume
            d['timestamp'] = datetime.now()
            
            if 'indicator' in d:
                result = d['indicator'].add_tick(ltp)
                if result:
                    for k, v in result.items():
                        d[k] = v
                
                if tick_count <= 10:
                    print(f"   {d['symbol']}: RSI={d.get('rsi', 0):.1f}")
                    
    except Exception as e:
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

# ====================================================================
# EXCEL FUNCTIONS
# ====================================================================

def setup_excel_headers():
    try:
        ws = excel_name.sheets['symbols']
        ws.range("1:1").clear_contents()
        
        headers = [
            'Symbol', 'LTP', 'Open', 'High', 'Low', 'Close', 'Volume',
            'RSI', 'StochRSI', 'SMA Stoch', 'BB Upper', 'BB Middle', 'BB Lower',
            'BUY', 'SELL', 'Signal', 'Last Update'
        ]
        
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.range((1, col_idx))
            cell.value = header
            
            if header in ['BUY', 'SELL', 'Signal']:
                cell.color = (255, 100, 100)
            else:
                cell.color = (54, 96, 146)
            cell.font.color = (255, 255, 255)
            cell.font.bold = True
        
        ws.range('A:A').column_width = 20
        ws.range('B:Q').column_width = 12
        
        return True
    except Exception as e:
        print(f"Error setting up headers: {e}")
        return False

def read_symbols_from_excel():
    try:
        ws = excel_name.sheets['symbols']
        symbols_data = ws.range("A2:A200").value
        symbols = []
        
        if symbols_data:
            for s in symbols_data:
                if s:
                    s_str = str(s).strip().upper()
                    if s_str.startswith('NSE:'):
                        s_str = s_str[4:]
                    if not s_str.endswith('-EQ'):
                        s_str = f"{s_str}-EQ"
                    symbols.append(s_str)
        
        return symbols
    except Exception:
        return []

def update_excel_bulk():
    global last_excel_update
    try:
        current_time = time.time()
        if current_time - last_excel_update < Config.EXCEL_UPDATE_INTERVAL:
            return
        last_excel_update = current_time
        
        ws = excel_name.sheets['symbols']
        
        # Read symbols from column A
        symbols_list = ws.range("A2:A200").value
        
        if not symbols_list:
            return
        
        # Prepare data for columns B to Q (16 columns total: B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q)
        # That's 16 columns from B to Q
        data_rows = []
        
        for symbol_val in symbols_list:
            if not symbol_val:
                # Empty row - 16 empty strings for columns B to Q
                data_rows.append([''] * 16)
                continue
            
            symbol = str(symbol_val).strip().upper()
            if symbol.startswith('NSE:'):
                symbol = symbol[4:]
            if not symbol.endswith('-EQ'):
                symbol = f"{symbol}-EQ"
            
            tk = symbol_tokens.get(symbol)
            
            if tk and tk in live_data:
                d = live_data[tk]
                # 16 columns: B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q
                data_rows.append([
                    d.get('ltp', ''),           # B
                    d.get('open', ''),          # C
                    d.get('high', ''),          # D
                    d.get('low', ''),           # E
                    d.get('close', ''),         # F
                    d.get('volume', ''),        # G
                    d.get('rsi', ''),           # H
                    d.get('stoch_rsi', ''),     # I
                    d.get('sma_stoch', ''),     # J
                    d.get('bb_upper', ''),      # K
                    d.get('bb_middle', ''),     # L
                    d.get('bb_lower', ''),      # M
                    d.get('buy', ''),           # N
                    d.get('sell', ''),          # O
                    d.get('signal', ''),        # P
                    d['timestamp'].strftime('%H:%M:%S') if d.get('timestamp') else ''  # Q
                ])
            else:
                # Symbol exists but not initialized yet - 16 empty strings
                data_rows.append([''] * 16)
        
        # Update ONLY columns B to Q (16 columns)
        if data_rows:
            ws.range(f"B2:Q{2 + len(data_rows) - 1}").value = data_rows
            
    except Exception as e:
        print(f"Excel update error: {e}")

def check_new_symbols():
    global last_symbol_check
    try:
        ws = excel_name.sheets['symbols']
        symbols_data = ws.range("A2:A100").value
        
        if not symbols_data:
            return
        
        # Get current symbols from Excel
        current_symbols = []
        for s in symbols_data:
            if s:
                s_str = str(s).strip().upper()
                if s_str.startswith('NSE:'):
                    s_str = s_str[4:]
                if not s_str.endswith('-EQ'):
                    s_str = f"{s_str}-EQ"
                current_symbols.append(s_str)
            else:
                current_symbols.append('')
        
        # Find new symbols (in Excel but not in symbol_tokens)
        new_symbols = []
        for symbol in current_symbols:
            if symbol and symbol not in symbol_tokens:
                new_symbols.append(symbol)
        
        if new_symbols:
            print(f"\n🆕 Found {len(new_symbols)} new symbols: {new_symbols}")
            new_tokens = []
            
            for symbol in new_symbols:
                try:
                    print(f"   Fetching data for {symbol}...")
                    token = GetToken("NSE", symbol)
                    
                    if token:
                        tk = f"NSE|{token}"
                        symbol_tokens[symbol] = tk
                        token_symbols[token] = symbol
                        
                        # Fetch historical data
                        hist_data = fetch_historical_data(symbol)
                        
                        # Create indicator
                        indicator = StochRSIIndicator(symbol, hist_data['all_closes'] if hist_data else [])
                        
                        # Initialize live data
                        live_data[tk] = {
                            'symbol': symbol, 
                            'first_tick': True,
                            'indicator': indicator,
                            'ltp': 0, 'volume': 0, 
                            'open': 0, 'high': 0, 'low': 0, 'close': 0,
                            'rsi': indicator.current_rsi if hasattr(indicator, 'current_rsi') else 50,
                            'stoch_rsi': indicator.current_stoch if hasattr(indicator, 'current_stoch') else 50,
                            'sma_stoch': indicator.current_sma_stoch if hasattr(indicator, 'current_sma_stoch') else 50,
                            'bb_upper': indicator.current_bb_upper if hasattr(indicator, 'current_bb_upper') else 0,
                            'bb_middle': indicator.current_bb_middle if hasattr(indicator, 'current_bb_middle') else 0,
                            'bb_lower': indicator.current_bb_lower if hasattr(indicator, 'current_bb_lower') else 0,
                            'buy': '', 'sell': '', 'signal': '', 
                            'timestamp': None
                        }
                        
                        new_tokens.append(tk)
                        
                        # Store historical data
                        if hist_data:
                            historical_data_cache[symbol] = hist_data['yesterday']
                            print(f"   ✅ Added {symbol} with {len(hist_data['all_closes'])} days of data")
                        else:
                            print(f"   ✅ Added {symbol} (no historical data)")
                            
                    else:
                        print(f"   ❌ Could not get token for {symbol}")
                        
                except Exception as e:
                    print(f"   ❌ Error adding {symbol}: {e}")
                    
                time.sleep(0.05)
            
            # Subscribe to new tokens
            if new_tokens and feed_opened:
                subscribe_symbols(new_tokens)
                print(f"✓ Subscribed to {len(new_tokens)} new symbols\n")
                
            # Force Excel update
            update_excel_bulk()
            
    except Exception as e:
        print(f"Error in check_new_symbols: {e}")

def start_excel_loop():
    global last_symbol_check, tick_count
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
            
        except Exception:
            time.sleep(0.1)

# ====================================================================
# MAIN FUNCTION
# ====================================================================

def main():
    global historical_data_cache
    
    print("\n" + "="*80)
    print("🚀 STOCHRSI + BOLLINGER BANDS TRADING SYSTEM")
    print("="*80)
    
    print("\n[1/4] Setting up Excel...")
    setup_excel_headers()
    
    print("\n[2/4] Logging to Shoonya...")
    if not Shoonya_login():
        print("❌ Login failed! Check your LOGIN sheet in Excel.")
        return
    
    print("\n[3/4] Reading symbols from Excel...")
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
    print(f"\n[4/4] Loading {Config.LOAD_DAYS} days of historical data...")
    
    for i, symbol in enumerate(symbols, 1):
        print(f"   [{i:2}/{len(symbols)}] {symbol}...", end=" ")
        token = GetToken("NSE", symbol)
        
        if token:
            tk = f"NSE|{token}"
            symbol_tokens[symbol] = tk
            token_symbols[token] = symbol
            
            # Fetch historical data
            hist_data = fetch_historical_data(symbol)
            
            # Create indicator
            indicator = StochRSIIndicator(symbol, hist_data['all_closes'] if hist_data else [])
            
            # Initialize live data
            live_data[tk] = {
                'symbol': symbol, 
                'first_tick': True,
                'indicator': indicator,
                'ltp': 0, 'volume': 0, 
                'open': 0, 'high': 0, 'low': 0, 'close': 0,
                'rsi': indicator.current_rsi if hasattr(indicator, 'current_rsi') else 50,
                'stoch_rsi': indicator.current_stoch if hasattr(indicator, 'current_stoch') else 50,
                'sma_stoch': indicator.current_sma_stoch if hasattr(indicator, 'current_sma_stoch') else 50,
                'bb_upper': indicator.current_bb_upper if hasattr(indicator, 'current_bb_upper') else 0,
                'bb_middle': indicator.current_bb_middle if hasattr(indicator, 'current_bb_middle') else 0,
                'bb_lower': indicator.current_bb_lower if hasattr(indicator, 'current_bb_lower') else 0,
                'buy': '', 'sell': '', 'signal': '', 
                'timestamp': None
            }
            
            if hist_data:
                historical_data_cache[symbol] = hist_data['yesterday']
                print(f"✓ ({len(hist_data['all_closes'])} days)")
            else:
                print(f"✓ (No historical data)")
        else:
            print(f"✗ FAILED - Token not found")
        
        time.sleep(0.03)
    
    print(f"\n✅ Initialized {len(symbol_tokens)} symbols")
    
    if len(symbol_tokens) == 0:
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
            if symbol_tokens:
                print(f"\n📡 Subscribing to {len(symbol_tokens)} symbols...")
                subscribe_symbols(list(symbol_tokens.values()))
                print(f"✓ Subscribed to all symbols")
            
            print("\n🚀 StochRSI Trading System Running...")
            print("   ✅ RSI = Wilder's RMA (TradingView exact)")
            print("   ✅ StochRSI = ta.stoch() on RSI")
            print("   ✅ BUY/SELL signals active")
            print("   ✅ Historical data loaded\n")
            
            start_excel_loop()
        else:
            print("❌ WebSocket connection failed after 15 seconds!")
            
    except Exception as e:
        print(f"❌ Failed to start WebSocket: {e}")

if __name__ == "__main__":
    main()
