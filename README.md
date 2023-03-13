# <h1 align="center">xlsb_converter</h1>

**Simple python script that converts all Excel files (xls, xlsx, xlsm, csv) in a directory into xlsb files.**

[![Test OS](https://img.shields.io/badge/runs_on-windows-blue.svg)](https://github.com/gibz104/google-sheets-writer/actions/workflows/tests.yaml)

# Background

Do you have massive excel files for financial modeling?  Have you heard that storing your Excel files as binary format (.xlsb) can speed up your workbook's performance?  Xlsb files also support macros, similar to xlsm.  Well if you have hundreds of Excel files that need to be converted, then this script is for you.

# Usage

Simply pass a windows directory path to the 'main' function and all xls, xlsx, xlsm, and csv files in that directory will be saved as xlsb files in the same location.  

This script does not delete the original files, but rather saves a new version of the originals.  If the script comes across a password protected workbook, you will be prompted to enter that workbook's password in an Excel popup window for the script to proceed.

# Disclaimer

This repository uses `WIN32COM`, which makes this script only work on Windows.
