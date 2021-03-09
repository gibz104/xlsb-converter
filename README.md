# xlsb_converter
Simple python script that converts all Excel files (xls, xlsx, xlsm, csv) in a directory into xlsb files.

***ONLY SUPPORTS WINDOWS***

Simply pass a windows directory path to the 'main' function and all xls, xlsx, xlsm, and csv files in that directory will be saved as xlsb files in the same location.  This script does not delete the original files, but rather saves a new version of the originals.  This repository uses WIN32COM, which is why this script only supports Windows.  If the script comes across a password protected workbook, you will be prompted to enter that workbook's password in an Excel popup window for the script to proceed.

***EXAMPLE:***

![image](https://user-images.githubusercontent.com/35471104/110549713-1ce5d080-80f8-11eb-8eec-455ae604d7ea.png)

