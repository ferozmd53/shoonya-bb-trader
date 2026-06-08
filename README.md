# Shoonya Trading System

## Installation

```bash
pip install -e .
```

## Usage

### 1. Create Excel file `symbols.xlsx` with sheet `LOGIN`

### 2. Get Auth Code

```bash
shoonya-auth
```

### 3. Run Trading System

```python
from shoonya_bb import BollingerBandsTrader
trader = BollingerBandsTrader()
trader.run()
```
