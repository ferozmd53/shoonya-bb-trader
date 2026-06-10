# Shoonya REAL TIME TICK with MULTI SYMBOLS  HISTORICAL DATA 

## Installation  and CODE

```bash  C M D
pip uninstall Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators -y
pip install git+https://github.com/ferozmd53/Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators.git

# run.py - Minimal CODE
#excel file download
from get_auth import get_auth_code
from Extreme_Reversal_Signal import main

get_auth_code()
main()

# ====================================================================
# DISCLAIMER
# ====================================================================
# This software is for educational and research purposes only.
# RISK WARNING:
# - Trading in financial markets involves substantial risk of loss.
# - Past performance does not guarantee future results.
# - This system generates trading signals based on technical indicators,
#   which may produce false signals.
# - You should consult with a qualified financial advisor before making
#   any trading decisions.
# - The author assumes no responsibility for any financial losses
#   incurred while using this software.
# - Use this software at your own risk.
#
# LIMITATIONS:
# - Real-time data depends on API availability and network conditions.
# - Historical data may vary between different data providers.
# - Indicators are calculated to match TradingView but may have
#   minor variations due to data source differences.
#
# By using this software, you agree that:
# 1. You understand the risks involved in trading.
# 2. You will test thoroughly in paper trading before live use.
# 3. You will not hold the author liable for any losses.
# ====================================================================
