# setup.py
from setuptools import setup, find_packages

setup(
    name="shoonya-bb-trader",
    version="1.0.0",
    author="ferozmd53",
    description="Bollinger Bands Trading System for Shoonya API",
    long_description="Real-time Bollinger Bands trading system with Excel integration",
    long_description_content_type="text/markdown",
    url="https://github.com/ferozmd53/shoonya-bb-trader",
    packages=find_packages(),
    py_modules=["bb_trader", "get_auth"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
    ],
    python_requires=">=3.8",
    install_requires=[
        "NorenRestApiPy>=0.0.22",
        "xlwings>=0.30.0",
        "pandas>=2.0.0",
        "numpy>=1.24.0",
        "selenium>=4.15.0",
        "webdriver-manager>=4.0.0",
        "pyotp>=2.9.0",
        "requests>=2.31.0",
    ],
    entry_points={
        "console_scripts": [
            "bb-trader=bb_trader:main",
            "bb-auth=get_auth:get_auth_code",
        ],
    },
)
