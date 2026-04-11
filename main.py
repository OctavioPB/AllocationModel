"""
main.py
-------
Entry point for the AllocationModel desktop application.

Usage:
    python main.py

Double-click this file (or run from the terminal) to launch the GUI.
Python 3.9+ and the packages in requirements.txt must be installed.
"""

import sys
import os

# Ensure the project root is on the import path regardless of how the
# script is invoked (double-click, terminal, IDE).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.gui import run

if __name__ == "__main__":
    run()
