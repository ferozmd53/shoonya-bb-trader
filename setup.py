from setuptools import setup, find_packages

setup(
    name="shoonya-trader",
    version="1.0.0",
    author="Your Name",
    description="Shoonya API Bollinger Bands Trading System",
    packages=find_packages(),
    python_requires=">=3.8",
    install_requires=[
        "NorenRestApiPy",
        "xlwings",
        "pandas",
        "numpy",
        "selenium",
        "webdriver-manager",
        "pyotp",
    ],
    entry_points={
        "console_scripts": [
            "shoonya-auth=shoonya_oauth.auth:main",
        ],
    },
)
