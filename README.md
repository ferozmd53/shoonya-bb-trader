# Shoonya REAL TIME ALL HISTORICAL DATA WITH TICK WITH Bollinger Bands Trading System

## Installation

```bash
pip uninstall shoonya-bb-trader -y
pip install git+https://github.com/ferozmd53/shoonya-bb-trader.git

# run.py - Minimal CODE
from get_auth import get_auth_code
from bb_trader import main

get_auth_code()
main()
