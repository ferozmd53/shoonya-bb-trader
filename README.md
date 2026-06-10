# Shoonya StochRSI Trading System
# DISCLAIMER
# This software is for educational and research purposes only.
## Installation
```bash
pip uninstall Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators -y
pip install git+https://github.com/ferozmd53/Shoonya-Multi-Symbol-Live-Historical-Data-with-Indicators.git
pip install NorenRestApiOAuth
pip install NorenRestApiPy 

# run.py - Minimal code
from get_auth import get_auth_code
from Extreme_Reversal_Signal import main

if __name__ == "__main__":
    get_auth_code()
    main()
