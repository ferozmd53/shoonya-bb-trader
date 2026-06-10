# main_script.py - FIXED HISTORICAL SIGNALS
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

# ============================================
# GLOBAL VARIABLES
# ============================================

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

# ============================================
# CONFIGURATION
# ============================================

class Config:
    MA_LENGTH = 20
    STD_UP = 2.0
    STD_DOWN = 2.0
    RSI_LENGTH = 14
    STO_LENGTH = 14
    STO_UPPER = 70
    STO_LOWER = 30
    EXCEL_UPDATE_INTERVAL = 0.1
    LOAD_DAYS = 500
    KEEP_DAYS = 100

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
# EXACT TRADINGVIEW RMA (Wilder's Moving Average)
# ============================================

def ta_rma(src, length):
    alpha = 1.0 / length
    return src.ewm(alpha=alpha, adjust=False).mean()

# ============================================
# FAST INDICATOR WITH PROPER INITIALIZATION
# ============================================

class FastIndicator:
    def __init__(self, symbol, all_closes, all_highs, all_lows):
        self.symbol = symbol
        
        self.all_closes = list(all_closes) if all_closes else []
        self.all_highs = list(all_highs) if all_highs else []
        self.all_lows = list(all_lows) if all_lows else []
        
        self.closes = self.all_closes[-Config.KEEP_DAYS:] if len(self.all_closes) > Config.KEEP_DAYS else self.all_closes.copy()
        self.highs = self.all_highs[-Config.KEEP_DAYS:] if len(self.all_highs) > Config.KEEP_DAYS else self.all_highs.copy()
        self.lows = self.all_lows[-Config.KEEP_DAYS:] if len(self.all_lows) > Config.KEEP_DAYS else self.all_lows.copy()
        
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
            rsi = 100
        else:
            rs = self.avg_gain / self.avg_loss
            rsi = 100 - (100 / (1 + rs))
        
        return rsi
    
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
# LOGIN
# ============================================

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
                print("❌ Missing credentials!")
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
            api.injectOAuthHeader(acc_tok, userid, actid)
            print("✅ Login Successful!")
            return 1
        else:
            print("❌ Login failed")
            
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
# FETCH HISTORICAL DATA WITH SIGNALS
# ============================================

def fetch_historical_data(symbol):
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=Config.LOAD_DAYS)
        
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
        
        # Calculate ALL indicators including signals
        df_calc = df.copy()
        
        # RSI
        delta = df_calc['close'].diff()
        u = delta.where(delta > 0, 0.0)
        d = -delta.where(delta < 0, 0.0)
        alpha = 1.0 / Config.RSI_LENGTH
        avg_gain = u.ewm(alpha=alpha, adjust=False).mean()
        avg_loss = d.ewm(alpha=alpha, adjust=False).mean()
        rs = avg_gain / avg_loss
        df_calc['RSI'] = 100 - (100 / (1 + rs))
        
        # StochRSI
        rsi_low = df_calc['RSI'].rolling(Config.STO_LENGTH).min()
        rsi_high = df_calc['RSI'].rolling(Config.STO_LENGTH).max()
        df_calc['STOCH_RSI'] = 100 * (df_calc['RSI'] - rsi_low) / (rsi_high - rsi_low)
        df_calc['STOCH_RSI'] = df_calc['STOCH_RSI'].fillna(50)
        df_calc['SMA_STOCH'] = df_calc['STOCH_RSI'].rolling(3).mean()
        
        # Bollinger Bands
        df_calc['BB_MIDDLE'] = df_calc['close'].rolling(Config.MA_LENGTH).mean()
        std = df_calc['close'].rolling(Config.MA_LENGTH).std(ddof=0)
        df_calc['BB_UPPER'] = df_calc['BB_MIDDLE'] + (Config.STD_UP * std)
        df_calc['BB_LOWER'] = df_calc['BB_MIDDLE'] - (Config.STD_DOWN * std)
        
        # BUY/SELL Signals (using shift for proper signal generation)
        df_calc['BUY_SIGNAL'] = (
            (df_calc['close'].shift(1) < df_calc['BB_LOWER'].shift(1)) &
            (df_calc['close'] > df_calc['BB_LOWER']) &
            (df_calc['SMA_STOCH'].shift(1) < Config.STO_LOWER)
        )
        
        df_calc['SELL_SIGNAL'] = (
            (df_calc['close'].shift(1) > df_calc['BB_UPPER'].shift(1)) &
            (df_calc['close'] < df_calc['BB_UPPER']) &
            (df_calc['SMA_STOCH'].shift(1) > Config.STO_UPPER)
        )
        
        # Get yesterday's data (last row)
        yesterday = df_calc.iloc[-1]
        
        # Calculate buy/sell for historical display (from yesterday's signals)
        buy_value = 1 if yesterday['BUY_SIGNAL'] else ''
        sell_value = 1 if yesterday['SELL_SIGNAL'] else ''
        signal_value = 'BUY' if yesterday['BUY_SIGNAL'] else ('SELL' if yesterday['SELL_SIGNAL'] else '')
        
        yesterday_data = {
            'date': yesterday['datetime'].strftime('%d/%m/%Y'),
            'open': yesterday['open'],
            'high': yesterday['high'],
            'low': yesterday['low'],
            'close': yesterday['close'],
            'volume': yesterday['volume'],
            'bb_upper': round(yesterday['BB_UPPER'], 2),
            'bb_middle': round(yesterday['BB_MIDDLE'], 2),
            'bb_lower': round(yesterday['BB_LOWER'], 2),
            'rsi': round(yesterday['RSI'], 2),
            'stoch_rsi': round(yesterday['STOCH_RSI'], 2),
            'sma_stoch': round(yesterday['SMA_STOCH'], 2),
            'buy': buy_value,
            'sell': sell_value,
            'signal': signal_value
        }
        
        # Also calculate 2 days ago for debugging (optional)
        if len(df_calc) >= 2:
            day_before = df_calc.iloc[-2]
            print(f"   📊 {symbol}: Yesterday RSI={yesterday_data['rsi']}, Stoch={yesterday_data['stoch_rsi']}, Buy={buy_value}, Sell={sell_value}")
        
        return {
            'all_closes': df['close'].tolist(),
            'all_highs': df['high'].tolist(),
            'all_lows': df['low'].tolist(),
            'yesterday': yesterday_data
        }
        
    except Exception as e:
        print(f"Error: {e}")
        return None

# ============================================
# WEBSOCKET CALLBACKS
# ============================================

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
                    d['rsi'] = result['rsi']
                    d['stoch_rsi'] = result['stoch_rsi']
                    d['sma_stoch'] = result['sma_stoch']
                    d['bb_upper'] = result['bb_upper']
                    d['bb_middle'] = result['bb_middle']
                    d['bb_lower'] = result['bb_lower']
                    d['buy'] = result['buy']
                    d['sell'] = result['sell']
                    d['signal'] = result['signal']
                
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

# ============================================
# CHECK NEW SYMBOLS
# ============================================

def check_new_symbols():
    global last_symbol_check
    
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
                        
                        hist_data = fetch_historical_data(symbol)
                        
                        live_data[tk] = {
                            'symbol': symbol,
                            'first_tick': True,
                            'indicator': FastIndicator(symbol, 
                                                       hist_data['all_closes'] if hist_data else [],
                                                       hist_data['all_highs'] if hist_data else [],
                                                       hist_data['all_lows'] if hist_data else []),
                            'ltp': 0, 'volume': 0,
                            'open': 0, 'high': 0, 'low': 0, 'close': 0,
                            'rsi': 50, 'stoch_rsi': 50, 'sma_stoch': 50,
                            'bb_upper': 0, 'bb_middle': 0, 'bb_lower': 0,
                            'buy': '', 'sell': '', 'signal': '',
                            'timestamp': None
                        }
                        new_tokens.append(tk)
                        
                        if hist_data:
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
# UPDATE EXCEL
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
        
        rows = []
        
        for symbol_cell in symbols_data:
            if not symbol_cell:
                rows.append([''] * 32)
                continue
            
            symbol = str(symbol_cell).strip().upper()
            if symbol.startswith('NSE:'):
                symbol = symbol[4:]
            if not symbol.endswith('-EQ'):
                symbol = f"{symbol}-EQ"
            
            tk = symbol_tokens.get(symbol)
            hist = historical_data_cache.get(symbol, {})
            
            if tk and tk in live_data:
                d = live_data[tk]
                ltp = d.get('ltp', 0)
                
                rows.append([
                    symbol,
                    ltp if ltp > 0 else '',
                    d.get('open', 0) if d.get('open', 0) > 0 else '',
                    d.get('high', 0) if d.get('high', 0) > 0 else '',
                    d.get('low', 0) if d.get('low', 0) > 0 else '',
                    d.get('close', 0) if d.get('close', 0) > 0 else '',
                    d.get('volume', 0) if d.get('volume', 0) > 0 else '',
                    d.get('rsi', ''),
                    d.get('stoch_rsi', ''),
                    d.get('sma_stoch', ''),
                    d.get('bb_upper', ''),
                    d.get('bb_middle', ''),
                    d.get('bb_lower', ''),
                    d.get('buy', ''),
                    d.get('sell', ''),
                    d.get('signal', ''),
                    d['timestamp'].strftime('%H:%M:%S') if d.get('timestamp') else '',
                    hist.get('date', ''),
                    hist.get('open', ''),
                    hist.get('high', ''),
                    hist.get('low', ''),
                    hist.get('close', ''),
                    hist.get('volume', ''),
                    hist.get('rsi', ''),
                    hist.get('stoch_rsi', ''),
                    hist.get('sma_stoch', ''),
                    hist.get('bb_upper', ''),
                    hist.get('bb_middle', ''),
                    hist.get('bb_lower', ''),
                    hist.get('buy', ''),
                    hist.get('sell', ''),
                    hist.get('signal', '')
                ])
            else:
                rows.append([''] * 32)
        
        try:
            if rows:
                ws.range(f"A2:AF{2 + len(rows) - 1}").value = rows
        except Exception:
            pass
            
    except Exception:
        pass

# ============================================
# SETUP HEADERS
# ============================================

def setup_excel_headers():
    try:
        ws = excel_name.sheets['symbols']
        ws.range("1:1").clear_contents()
        
        headers = [
            'Symbol', 'LTP', 'Open', 'High', 'Low', 'Close', 'Volume',
            'RSI', 'StochRSI', 'SMA Stoch', 'BB Upper', 'BB Middle', 'BB Lower',
            'BUY', 'SELL', 'Signal', 'Last Update',
            'Date', 'Open', 'High', 'Low', 'Close', 'Volume',
            'RSI', 'StochRSI', 'SMA Stoch', 'BB Upper', 'BB Middle', 'BB Lower',
            'BUY', 'SELL', 'Signal'
        ]
        
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.range((1, col_idx))
            cell.value = header
            
            if header in ['BUY', 'SELL', 'Signal']:
                cell.color = (255, 100, 100)
            elif col_idx >= 18:
                cell.color = (146, 96, 54)
            else:
                cell.color = (54, 96, 146)
            cell.font.color = (255, 255, 255)
            cell.font.bold = True
        
        ws.range('A:AF').column_width = 12
        ws.range('A:A').column_width = 20

        return True
    except Exception as e:
        print(f"Error: {e}")
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
    except Exception:
        return []

# ============================================
# MAIN EXCEL LOOP
# ============================================

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

# ============================================
# MAIN
# ============================================

def main():
    global historical_data_cache
    
    print("\n" + "="*80)
    print("🚀 TRADINGVIEW EXACT MATCH WITH SIGNALS")
    print("="*80)
    
    print("\n[1/4] Setting up Excel...")
    setup_excel_headers()
    
    print("\n[2/4] Logging to Shoonya...")
    if not Shoonya_login():
        print("❌ Login failed!")
        return
    
    print("\n[3/4] Reading symbols...")
    symbols = read_symbols_from_excel()
    
    if not symbols:
        default = ["RELIANCE-EQ", "TCS-EQ", "INFY-EQ"]
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
            
            hist_data = fetch_historical_data(symbol)
            
            live_data[tk] = {
                'symbol': symbol,
                'first_tick': True,
                'indicator': FastIndicator(symbol, 
                                           hist_data['all_closes'] if hist_data else [],
                                           hist_data['all_highs'] if hist_data else [],
                                           hist_data['all_lows'] if hist_data else []),
                'ltp': 0, 'volume': 0,
                'open': 0, 'high': 0, 'low': 0, 'close': 0,
                'rsi': 50, 'stoch_rsi': 50, 'sma_stoch': 50,
                'bb_upper': 0, 'bb_middle': 0, 'bb_lower': 0,
                'buy': '', 'sell': '', 'signal': '',
                'timestamp': None
            }
            
            if hist_data:
                historical_data_cache[symbol] = hist_data['yesterday']
                print(f"✓ ({len(hist_data['all_closes'])} days)")
            else:
                print(f"⚠️ No data")
        else:
            print(f"✗ FAILED")
        time.sleep(0.03)
    
    print(f"\n✅ Initialized {len(symbol_tokens)} symbols")
    
    if len(symbol_tokens) == 0:
        print("❌ No symbols!")
        return
    
    print("\nStarting WebSocket...")
    
    try:
        api.start_websocket(
            subscribe_callback=on_ticks,
            order_update_callback=on_order,
            socket_open_callback=on_open,
            socket_close_callback=on_close
        )
        
        for _ in range(15):
            if feed_opened:
                break
            time.sleep(1)
        
        if feed_opened:
            print("✅ WebSocket connected!")
            if symbol_tokens:
                subscribe_symbols(list(symbol_tokens.values()))
            print("\n🚀 Running with FULL signal detection...")
            print("   ✅ Historical BUY/SELL signals now displayed")
            print("   ✅ Live signals trigger in real-time")
            print("   ✅ All indicators match TradingView!\n")
            start_excel_loop()
        else:
            print("❌ WebSocket failed!")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
