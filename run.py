"""
Launcher for X2T - Excel to Tally Converter
"""

import sys
import os

# Add x2t_desktop to sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(BASE_DIR, "x2t_desktop")
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from app import main

if __name__ == "__main__":
    main()
